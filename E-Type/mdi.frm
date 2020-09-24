VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.MDIForm frmMDI 
   BackColor       =   &H8000000C&
   Caption         =   "E-Type"
   ClientHeight    =   8145
   ClientLeft      =   1875
   ClientTop       =   2250
   ClientWidth     =   9795
   Icon            =   "mdi.frx":0000
   LinkTopic       =   "MDIForm1"
   OLEDropMode     =   1  'Manual
   Begin MSComctlLib.ImageList imlToolbHot 
      Left            =   9240
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   23
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":0E42
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":119A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":14F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":184A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":1BA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":1EFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":2252
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":27A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":2AFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":2E56
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":31AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":34FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":3A56
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":3FAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":4506
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":485A
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":4BB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":4F0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":5262
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":55BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":5912
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":5C6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":5FC2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlDriveFileList2 
      Left            =   9240
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":60D6
            Key             =   "Folder"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":63FA
            Key             =   "FolderOpen"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":674E
            Key             =   "FolderClosed"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":6AA2
            Key             =   "TextFile"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":6BB6
            Key             =   "UpDir"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":6F0A
            Key             =   "FolderO"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picSideBar 
      Align           =   3  'Align Left
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   7110
      Left            =   0
      ScaleHeight     =   7110
      ScaleWidth      =   2850
      TabIndex        =   2
      Tag             =   "big"
      Top             =   780
      Width           =   2850
      Begin VB.PictureBox picContainer 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   6735
         Left            =   0
         ScaleHeight     =   6735
         ScaleWidth      =   7170
         TabIndex        =   3
         Top             =   0
         Width           =   7170
         Begin VB.FileListBox File1 
            Appearance      =   0  'Flat
            Height          =   615
            Hidden          =   -1  'True
            Left            =   3360
            System          =   -1  'True
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   1080
            Width           =   2655
         End
         Begin VB.DirListBox Dir1 
            Appearance      =   0  'Flat
            Height          =   765
            Left            =   3360
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   120
            Width           =   2655
         End
         Begin TabDlg.SSTab sstDriveFilelist 
            Height          =   6135
            Left            =   0
            TabIndex        =   7
            Top             =   270
            Visible         =   0   'False
            Width           =   2805
            _ExtentX        =   4948
            _ExtentY        =   10821
            _Version        =   393216
            TabOrientation  =   3
            Style           =   1
            Tabs            =   4
            Tab             =   2
            TabsPerRow      =   4
            TabHeight       =   520
            WordWrap        =   0   'False
            ShowFocusRect   =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   " "
            TabPicture(0)   =   "mdi.frx":722E
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "lstLibraries"
            Tab(0).Control(1)=   "cmbLibraries"
            Tab(0).ControlCount=   2
            TabCaption(1)   =   " "
            TabPicture(1)   =   "mdi.frx":75C0
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "TVProjects"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   " "
            TabPicture(2)   =   "mdi.frx":7B5A
            Tab(2).ControlEnabled=   -1  'True
            Tab(2).Control(0)=   "ListView1"
            Tab(2).Control(0).Enabled=   0   'False
            Tab(2).Control(1)=   "Drive1"
            Tab(2).Control(1).Enabled=   0   'False
            Tab(2).Control(2)=   "cboFiletypes"
            Tab(2).Control(2).Enabled=   0   'False
            Tab(2).Control(3)=   "cboPath"
            Tab(2).Control(3).Enabled=   0   'False
            Tab(2).Control(4)=   "cmdPathMenu"
            Tab(2).Control(4).Enabled=   0   'False
            Tab(2).ControlCount=   5
            TabCaption(3)   =   " "
            TabPicture(3)   =   "mdi.frx":7EEC
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "lvFindResults"
            Tab(3).ControlCount=   1
            Begin VB.CommandButton cmdPathMenu 
               Caption         =   ">"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   260
               Left            =   2160
               TabIndex        =   19
               ToolTipText     =   "History"
               Top             =   490
               Width           =   240
            End
            Begin VB.ComboBox cboPath 
               Height          =   315
               ItemData        =   "mdi.frx":827E
               Left            =   30
               List            =   "mdi.frx":8280
               Style           =   1  'Simple Combo
               TabIndex        =   20
               Text            =   "Combo1"
               Top             =   465
               Width           =   2100
            End
            Begin VB.ComboBox cboFiletypes 
               Height          =   315
               ItemData        =   "mdi.frx":8282
               Left            =   30
               List            =   "mdi.frx":8284
               TabIndex        =   16
               Top             =   5760
               Width           =   2400
            End
            Begin VB.DriveListBox Drive1 
               Height          =   315
               Left            =   30
               TabIndex        =   10
               TabStop         =   0   'False
               Top             =   120
               Width           =   2400
            End
            Begin VB.ListBox lstLibraries 
               Height          =   5325
               Left            =   -74975
               TabIndex        =   9
               Top             =   480
               Width           =   2400
            End
            Begin VB.ComboBox cmbLibraries 
               Height          =   315
               Left            =   -74950
               TabIndex        =   8
               Text            =   "Combo2"
               Top             =   120
               Width           =   2400
            End
            Begin MSComctlLib.ListView ListView1 
               Height          =   4890
               Left            =   30
               TabIndex        =   11
               Top             =   800
               Width           =   2400
               _ExtentX        =   4233
               _ExtentY        =   8625
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               _Version        =   393217
               SmallIcons      =   "ImageList1"
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               Appearance      =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               NumItems        =   0
            End
            Begin MSComctlLib.TreeView TVProjects 
               Height          =   5835
               Left            =   -74970
               TabIndex        =   15
               Top             =   120
               Width           =   2415
               _ExtentX        =   4260
               _ExtentY        =   10292
               _Version        =   393217
               Indentation     =   353
               LineStyle       =   1
               Style           =   7
               ImageList       =   "imlDriveFileList2"
               Appearance      =   1
            End
            Begin MSComctlLib.ListView lvFindResults 
               Height          =   5955
               Left            =   -74970
               TabIndex        =   17
               Top             =   120
               Width           =   2400
               _ExtentX        =   4233
               _ExtentY        =   10504
               View            =   3
               MultiSelect     =   -1  'True
               LabelWrap       =   0   'False
               HideSelection   =   -1  'True
               AllowReorder    =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               SmallIcons      =   "ImageList1"
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               Appearance      =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               NumItems        =   0
            End
         End
         Begin VB.Line linLight 
            BorderColor     =   &H80000010&
            X1              =   2860
            X2              =   0
            Y1              =   0
            Y2              =   0
         End
         Begin VB.Line linDark 
            BorderColor     =   &H80000014&
            X1              =   2880
            X2              =   0
            Y1              =   10
            Y2              =   10
         End
         Begin VB.Label lblShowHide 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "<"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   195
            Left            =   2445
            TabIndex        =   14
            Top             =   45
            Width           =   120
         End
         Begin VB.Label lblTabInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Files"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   240
            Left            =   40
            TabIndex        =   13
            Top             =   45
            Width           =   1650
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   375
            Left            =   3360
            TabIndex        =   6
            Top             =   2160
            Width           =   2055
         End
         Begin VB.Label lblClose 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   195
            Left            =   2640
            TabIndex        =   12
            Top             =   60
            Width           =   135
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H80000010&
            BorderStyle     =   0  'Transparent
            FillColor       =   &H80000010&
            FillStyle       =   0  'Solid
            Height          =   290
            Left            =   20
            Top             =   0
            Width           =   2835
         End
      End
   End
   Begin MSComctlLib.StatusBar SBarMain 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   7890
      Width           =   9795
      _ExtentX        =   17277
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   2461
            MinWidth        =   2469
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   4948
            MinWidth        =   4940
            Text            =   "Line+Col"
            TextSave        =   "Line+Col"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3704
            MinWidth        =   3704
            Text            =   "Mod:"
            TextSave        =   "Mod:"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3175
            MinWidth        =   3175
            Text            =   "File Size: "
            TextSave        =   "File Size: "
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "CAPS"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9795
      _ExtentX        =   17277
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbNorm"
      HotImageList    =   "imlToolbHot"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   28
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New Document"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open Document"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save Document"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "saveAll"
            Object.ToolTipText     =   "Save All"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Separator"
            ImageIndex      =   23
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "Print"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "PrintPreview"
            Description     =   "5"
            Object.ToolTipText     =   "Print Priview"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   23
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   23
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Undo"
            Object.ToolTipText     =   "Undo"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Redo"
            Object.ToolTipText     =   "Redo"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   23
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Find"
            Object.ToolTipText     =   "Find"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "FindP"
            Object.ToolTipText     =   "Find Previous"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "FindN"
            Object.ToolTipText     =   "Find Next"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Replace"
            Object.ToolTipText     =   "Replace Text"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   23
            Style           =   3
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "NextWindow"
            Object.ToolTipText     =   "Activate Next Window"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "PrevWindow"
            Object.ToolTipText     =   "Activate Previous Window"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "TileH"
            Object.ToolTipText     =   "Tile Windows Horizontaly"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "TileV"
            Object.ToolTipText     =   "Tile Windows Verticaly"
            ImageIndex      =   19
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Cascade"
            Object.ToolTipText     =   "Cascade Windows"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "SelWindow"
            Object.ToolTipText     =   "Select Window"
            ImageIndex      =   21
            Style           =   5
         EndProperty
         BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   23
            Style           =   3
         EndProperty
         BeginProperty Button28 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Properties"
            Object.ToolTipText     =   "File Properties"
            ImageIndex      =   22
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CMDialog1 
      Left            =   9240
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "TXT"
      FilterIndex     =   557
      FontSize        =   1,27584e-37
   End
   Begin MSComctlLib.Toolbar tbaDocuments 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   18
      Top             =   360
      Width           =   9795
      _ExtentX        =   17277
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      ToolTips        =   0   'False
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
      EndProperty
      Begin MSComctlLib.TabStrip tabDocuments 
         Height          =   495
         Left            =   0
         TabIndex        =   21
         Top             =   50
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   873
         HotTracking     =   -1  'True
         Separators      =   -1  'True
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Doc Tab"
               Key             =   "key1"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imlToolbNorm 
      Left            =   9240
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   23
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":8286
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":87DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":8D36
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":928E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":97E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":9D3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":A296
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":A7EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":AD42
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":B29A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":B7EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":BD42
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":C296
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":C7EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":CD3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":D092
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":D5E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":DB3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":E08E
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":E5E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":EB36
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":F08A
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":F5E2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
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
   Begin VB.Menu mnuProjects 
      Caption         =   "Projects"
      Visible         =   0   'False
      Begin VB.Menu mnuAddDocument 
         Caption         =   "Add Document"
      End
      Begin VB.Menu mnuAddCurrentDocument 
         Caption         =   "Add Current Document"
      End
      Begin VB.Menu mnuRemoveDocument 
         Caption         =   "Remove Document"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNewGroup 
         Caption         =   "New Group"
      End
      Begin VB.Menu mnuDeleteGroup 
         Caption         =   "Delete Group"
      End
      Begin VB.Menu mnuRenameGroup 
         Caption         =   "Rename Group"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnusaveProjectsGroupNow 
         Caption         =   "Save Internal Project Group Now"
      End
      Begin VB.Menu mnusaveProjectsGroup 
         Caption         =   "Save Project Group as.."
      End
      Begin VB.Menu mnuLoadExsternalProjectGroup 
         Caption         =   "Load External Project Group"
      End
   End
   Begin VB.Menu mnuPathsHolder 
      Caption         =   "MnuPaths"
      Visible         =   0   'False
      Begin VB.Menu mnuPaths 
         Caption         =   "History"
         Index           =   0
      End
   End
   Begin VB.Menu mnuDocTabs 
      Caption         =   "DocTabs"
      Visible         =   0   'False
      Begin VB.Menu mnuDocProperties 
         Caption         =   "Doc Properties"
      End
      Begin VB.Menu mnuPrintDocument 
         Caption         =   "Print Document"
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
      Begin VB.Menu mnuCloseAll 
         Caption         =   "Close All"
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStyle 
         Caption         =   "Style"
         Begin VB.Menu mnuTabs 
            Caption         =   "Tabs"
         End
         Begin VB.Menu mnuButtons 
            Caption         =   "Buttons"
         End
         Begin VB.Menu mnuFlatButtons 
            Caption         =   "Flat Buttons"
         End
      End
      Begin VB.Menu mnuHideDocTabs 
         Caption         =   "Hide DocTabs"
      End
   End
End
Attribute VB_Name = "frmMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'* for items on picSideBar start
'Variables to workaround the fact that there is no ItemDblClick Event
Dim xpos As Long
Dim ypos As Long
Dim NameOfFile As String
Dim clmX As ColumnHeader
Dim Item As ListItem
Dim itmX As ListItem
Dim counter As Long
Dim Counter2 As Integer
Dim dname As String
Dim Fname As String
Dim TempDname As String
Dim CurrentDir As String
'* for items on picSideBar end
Dim pathMenuCount As Integer     ' variable to count pathmenuarray


Private Sub cboFiletypes_Click()
    ' Change pattern in listview using the
    ' Select Case statement with the ListIndex of the
    ' ComboBox control.
    With File1
        Select Case cboFiletypes.ListIndex
        Case 0
           .Pattern = "*.txt"
        Case 1
           .Pattern = "*.rtf"
        Case 2
           .Pattern = "*.bat"
        Case 3
           .Pattern = "*.ini"
        Case 4
           .Pattern = "*.sys"
        Case 5
           .Pattern = "*.*"
        End Select
    End With
        
    ListView1.ListItems.Clear 'Clear Out Old Items

'    lvFindResults.ListItems.Clear 'Clear Out Old Items

    'add file and dirnames to the listview
    PopulateListView

End Sub

Private Sub cboPath_Click()
    'Change the Directory itmes to equal the new Current Directory
    Dir1.Path = cboPath.text
    ListView1.ListItems.Clear 'Clear Out Old Items
    Drive1.Drive = Left(cboPath, 3)
    'add file and dirnames to the listview
    PopulateListView

End Sub

Private Sub cboPath_KeyPress(KeyAscii As Integer)
    ' if enter is pressed
    If KeyAscii = 13 Then
        'Change the Directory List Box to equal the new Current Directory
        Dir1.Path = cboPath.text
        ListView1.ListItems.Clear 'Clear Out Old Items
    
        'add file and dirnames to the listview
        PopulateListView
        KeyAscii = 0
    End If
End Sub

Private Sub cmdPathMenu_Click()
    Dim ypos As Integer
    ypos = cmdPathMenu.Top + 270
    If Me.tbaDocuments.Visible = True Then
        ypos = ypos + Me.tbaDocuments.Height
    End If
    If Me.tbToolBar.Visible = True Then
        ypos = ypos + Me.tbToolBar.Height
    End If
    PopupMenu mnuPathsHolder, vbPopupMenuLeftAlign, cmdPathMenu.Left + cmdPathMenu.Width, ypos

    picSideBar.SetFocus
End Sub


Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    On Error Resume Next
    'Change the Directory itmes to equal the new Current Directory
    ChDrive Drive1.Drive
    Dir1.Path = Drive1.Drive
    cboPath.text = (CurDir)
    ListView1.ListItems.Clear 'Clear Out Old Items

    'add file and dirnames to the listview
    PopulateListView

End Sub

Private Sub lblClose_Click()
    
    ' hiding the picSideBar
    picSideBar.Visible = False
    ' unchecking the mnuViewSidebar on all open documents.
'    Dim x                    'document array
'    'On Error Resume Next
'        For x = 1 To fIndex  ' number of open documents
'            Document(x).mnuViewSidebar.Checked = False
'        Next
    Dim vForm As Variant
    For Each vForm In Forms
        If Not TypeOf vForm Is MDIForm Then
            If vForm.MDIChild Then
                vForm.mnuViewSidebar.Checked = False
            End If
        End If
    Next    ' vform


    ' cascade all documents for good looks
    frmMDI.Arrange vbCascade

End Sub

Private Sub lblShowHide_Click()
    
    If picSideBar.Tag = "big" Then
        picSideBar.Width = 410
        picContainer.Left = -2440
        
        picSideBar.Tag = "small"
        lblShowHide.Caption = ">"
        ElseIf picSideBar.Tag = "small" Then
        picSideBar.Width = 2850
        picContainer.Left = 0
        lblShowHide.Caption = "<"
        picSideBar.Tag = "big"
    End If
    ' cascade all documents for good looks
    frmMDI.Arrange vbCascade

End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ' Sorting the listview acording to wich columsheader is clicked
    ListView1.SortKey = ColumnHeader.index - 1
End Sub

Private Sub ListView1_DblClick()
    
    On Error Resume Next

    If ListView1.HitTest(xpos, ypos) Is Nothing Then
        Exit Sub
    Else
        Set Item = ListView1.HitTest(xpos, ypos)
    End If
    
    'If you Click on a filename just exit this subroutine
    If Right(Dir1.Path, 1) <> "\" Then
        CurrentDir = Dir1.Path & "\"
    Else
        CurrentDir = Dir1.Path
    End If
    
    If (GetAttr(CurrentDir & Item) And vbDirectory) <= 0 Then ' not a folder
    Dim strOpenFileName As String
    strOpenFileName = (CurrentDir & Item)
    
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
                        Me.ActiveForm.Text1.SetFocus
                    End Select
                Else
                    MsgBox "Document already open and in the same state.", vbInformation
                    Me.ActiveForm.Text1.SetFocus
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
        If frmMDI.SBarMain.Visible = True Then
            GetFileStats    ' filedata, date,size ect.
        End If
        Exit Sub
    End If
    
    
    ListView1.ListItems.Clear 'Clear Out Old Items
    
    'Change to selected Directory - Let Visual Basic do the work
    ChDir Item
    
    'Change the Directory List Box to equal the new Current Directory
    Dir1.Path = CurDir
    cboPath.text = (CurDir)
    'add file and dirnames to the listview
    PopulateListView
    
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    xpos = x
    ypos = y

End Sub



Private Sub lvFindResults_DblClick()
    ' open selected document
    OpenFile lvFindResults.SelectedItem.SubItems(2) & lvFindResults.SelectedItem
    ' Update the list of recently opened files in the File menu control array.
    UpdateFileMenu (lvFindResults.SelectedItem.SubItems(2) & lvFindResults.SelectedItem)
    ' If statusbar is visible, update infomation on form
    If frmMDI.SBarMain.Visible = True Then
        GetFileStats    ' filedata, date,size ect.
    End If

End Sub

Private Sub MDIForm_Activate()
    On Error GoTo LocalErrorHandler
    
    ' Load toolbarsettings from registry
    tbToolBar.RestoreToolbar App.Title, "Toolbar", "Toolbar1"
    ' make shure the toolbar is placed on top of doctabs
    frmMDI.tbToolBar.Top = 0
LocalErrorHandler:
    Resume Next
End Sub

Private Sub MDIForm_Load()
    On Error GoTo LocalErrorHandler

    ' VB insists on having one tab on the tabstrip, therfore clear all
    ' existing tabsfirst.
    tabDocuments.Tabs.Clear

    ' setting fIndex to 1.
    fIndex = 1
    ' Get all settings from registry
    ' Reading the forms last position
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 8500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
    ' Apperance
    tbToolBar.Visible = GetSetting(App.Title, "Apperance", "ShowToolBar", 0)
    SBarMain.Visible = GetSetting(App.Title, "Apperance", "ShowStatusBar", 0)
    tbaDocuments.Visible = GetSetting(App.Title, "Apperance", "ShowDocumentTabs", 0)
    picSideBar.Visible = GetSetting(App.Title, "Apperance", "ShowSideBar", 0)

    ' setup print header and footer
    sPrintHeader = GetSetting(App.Title, "Print", "Header", "Document Name: ^N" & vbCrLf & "Printed on ^D ^T")
    sPrintFooter = GetSetting(App.Title, "Print", "Footer", "***** END DOCUMENT *****" & vbCrLf & "^P")
    gLeftMargin = GetSetting(App.Title, "Print", "gLeftMargin", 25)
    gRightMargin = GetSetting(App.Title, "Print", "gRightMargin", 25)
    gTopMargin = GetSetting(App.Title, "Print", "gTopMargin", 25)
    gBottomMargin = GetSetting(App.Title, "Print", "gBottomMargin", 25)

    ' Application starts here (Load event of Startup form).
'    Show
    ' Always set the working directory to the directory containing the application.
    ChDir App.Path
        
    ' Initialize the document form array, and show the first document.
    ReDim Document(1 To 1)  ' array starts on 1
    ReDim FState(1 To 1)
    Document(1).Tag = 1
    Document(1).Caption = "Untitled: " & Document(1).Tag

    ' Read System registry and set the recent menu file list control array appropriately.
    GetRecentFiles
    ' Set public variable gFindDirection which determines which direction
    ' the FindIt function will search in.
    gFindDirection = 1
    
    '  See if TipOfDay should be shown at startup
    If GetSetting(App.Title, "StartUp", "ShowTip", 1) = 1 Then
        ' if tip is to be shown, unload splash
        frmSplash.trUnloadSplash.Enabled = True
        Load frmTip
        frmTip.Show vbModeless, frmMDI ' Me makes the window minimized with the main form
    End If
    
    ' if PicSidebar is visible, populate the filelist with files
    If picSideBar.Visible = True Then
        Call StartSideBar
    End If
    
    ' If statusbar is visible, update infomation on form
    If frmMDI.SBarMain.Visible = True Then
        GetFileStats    ' filedata, date,size ect.
    End If
          
    ' last thing to do is unload splash
    frmSplash.trUnloadSplash.Enabled = True

    ' if program is loaded from commandline or started from association
    ' THIS MUST BE AT THE END OF SUB!
    If Command$ = "" Then Exit Sub
    If Command$ = " " Then Exit Sub
    ' Call the file open procedure, passing a
    ' reference to the selected file name
    OpenFile Command$
    ' Update the list of recently opened files in the File menu control array.
    UpdateFileMenu Command$

LocalErrorHandler:
    MsgBox "Error # " & "<" & CStr(Err.Number) & "> " & Chr(13) & _
    "Error Description: " & Err.Description & Chr(13) & _
    "Error Source:" & "[" & Err.Source & "]", vbCritical

End Sub

Private Sub MDIForm_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error GoTo LocalErrorHandler
    
    'Count number of files
    Dim numFiles As Integer
    numFiles = Data.Files.Count
    
    'Open all dropped files
    Dim i As Integer
    For i = 1 To numFiles
        'File or directory?
        If (GetAttr(Data.Files(i)) And vbDirectory) = vbDirectory Then
            Me.OLEDropMode = 0
        Else
            ' check if the file is already open
            Dim a As Integer
            Dim ArrayCount As Integer
        
            ArrayCount = UBound(Document)
        
            ' Cycle through the document array. If one of the
            ' documents filenames matches the file droped, inform
            ' the user that the file already is open.
            For a = 1 To ArrayCount
                If Document(a).Caption = Data.Files(i) Then
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
                                OpenFile Data.Files(i)
                                ' Update the list of recently opened files in the File menu control array.
                                UpdateFileMenu Command$
                                ' If statusbar is visible, update infomation on form
                                If frmMDI.SBarMain.Visible = True Then
                                    GetFileStats    ' filedata, date,size ect.
                                End If
                            Case 7      ' User chose No.
                                ' nothing
                                Me.ActiveForm.Text1.SetFocus
                            End Select
                    Else
                        MsgBox "Document already open and in the same state.", vbInformation
                        Me.ActiveForm.Text1.SetFocus
                    End If
                ' the document is already open an in the same state.
                ' No need to open it
                Exit Sub
                End If
            Next
            ' The file(s) droped is not already open.
            ' Call the file open procedure, passing a
            ' reference to the selected file name
            OpenFile Data.Files(i)
            ' Update the list of recently opened files in the File menu control array.
            UpdateFileMenu Command$
            ' If statusbar is visible, update infomation on form
            If frmMDI.SBarMain.Visible = True Then
                GetFileStats    ' filedata, date,size ect.
            End If
        End If
    Next i
    Exit Sub
    
LocalErrorHandler:
    MsgBox "Error # " & "<" & CStr(Err.Number) & "> " & Chr(13) & _
    "Error Description: " & Err.Description & Chr(13) & _
    "Error Source:" & "[" & Err.Source & "]", vbCritical

End Sub

Private Sub MDIForm_Resize()
    resizeME
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    
    ' Saving the forms position
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
    
    ' Saving the printersettings
    SaveSetting App.Title, "Print", "Header", sPrintHeader
    SaveSetting App.Title, "Print", "Footer", sPrintFooter
    SaveSetting App.Title, "Print", "gLeftMargin", gLeftMargin
    SaveSetting App.Title, "Print", "gRightMargin", gRightMargin
    SaveSetting App.Title, "Print", "gTopMargin", gTopMargin
    SaveSetting App.Title, "Print", "gBottomMargin", gBottomMargin


    'save the projects to file
    Call SaveTreeViewToFile((App.Path) & "\" & "projects.etp", TVProjects)

    ' If the Unload event was not cancelled (in the QueryUnload events for the Notepad forms),
    ' there will be no document window left, so go ahead and end the application.
    If Not AnyPadsLeft() Then
        End
    End If
End Sub

Private Sub mnuAddCurrentDocument_Click()
    If Left(Me.ActiveForm.Caption, 8) = "Untitled" Then
        MsgBox "File must be saved first.", vbInformation + vbOKOnly
    Else
        Dim Name As String
        Dim Person As Node
        Dim Group As Node
    
        Name = Me.ActiveForm.Caption
        
        ' Find the group that should hold the new node.
        ' if the selected item is a parent
        If TVProjects.SelectedItem.Image = "Folder" Then
            Set Group = TVProjects.SelectedItem
            Set Person = TVProjects.Nodes.Add(Group, tvwChild, , Name, "TextFile")
        Else ' if selected item is a child
            Set Group = TVProjects.SelectedItem.Parent
            Set Person = TVProjects.Nodes.Add(Group, tvwChild, , Name, "TextFile")
        End If
    
        Person.EnsureVisible
    End If

End Sub

Private Sub mnuAddDocument_Click()
    'Call AddDocToGroup(TVProjects)
     frmFavorites.Show vbModeless, frmMDI ' Me makes the window minimized with the main form

End Sub

Private Sub mnuButtons_Click()
    Me.tabDocuments.Style = tabButtons
End Sub

Private Sub mnuClose_Click()
    ' Unload the corresponding form.
    '**** Her burdte programmet finne ut av hvilken tab musen er over
    Dim dNum As String
    dNum = (Mid(tabDocuments.SelectedItem.key, 4))
    Unload Document(dNum)

End Sub

Private Sub mnuCloseAll_Click()
    Dim x
    
    For x = 1 To UBound(Document)     ' number of open files
        Unload Document(x)  ' unload all open files
    Next

End Sub

Private Sub mnuDeleteGroup_Click()
    

'    If TVProjects.SelectedItem Is TVProjects.SelectedItem.Root Then
'        MsgBox "You can not delete the Root project!", vbOKOnly + vbExclamation
'        Exit Sub
'    End If
    
        Dim msg, Style, Title, Help, Ctxt, Response, MyString
    msg = "Remove " & TVProjects.SelectedItem.text & " and all" & Chr(13) & "documents in project group?"  ' Define message.
    Style = vbYesNo + vbQuestion + vbDefaultButton2   ' Define buttons.
    Title = "Warning"   ' Define title.
    ' Display message.
    Response = MsgBox(msg, Style, Title, Help, Ctxt)
    If Response = vbYes Then   ' User chose Yes.
       TVProjects.Nodes.Remove TVProjects.SelectedItem.index
    Else   ' User chose No.
       'nothing
    End If
    
End Sub

Private Sub mnuDocProperties_Click()
    frmFileInfo.Show
End Sub

Private Sub mnuFileExit_Click()
    ' End the application.
    End
End Sub

Private Sub mnuFileNew_Click()
    On Error GoTo LocalErrorHandler
    ' Call the new file procedure
    Call FileNew
    If SBarMain.Visible = True Then
        'Get the filestats from the current active file
        GetFileStats
    End If
    Exit Sub
    
LocalErrorHandler:
    MsgBox "Error # " & "<" & CStr(Err.Number) & "> " & Chr(13) & _
    "Error Description: " & Err.Description & Chr(13) & _
    "Error Source:" & "[" & Err.Source & "]", vbCritical
End Sub

Private Sub mnuFileOpen_Click()
    ' Call the file open procedure.
    FileOpenProc
    ' If statusbar is visible, update infomation on form
    If frmMDI.SBarMain.Visible = True Then
        GetFileStats    ' filedata, date,size ect.
    End If
End Sub

Private Sub mnuFlatButtons_Click()
    Me.tabDocuments.Style = tabFlatButtons
End Sub

Private Sub mnuHideDocTabs_Click()
    Dim x
    
    frmMDI.tbaDocuments.Visible = False
    ' cycle through all documents and uncheck the menues
    For x = 1 To UBound(Document)
        Document(x).mnuViewDocTabs.Checked = False
    Next x
    ' resize the controls on frmMDI.sidebar
    frmMDI.resizeME

End Sub

Private Sub mnuLoadExsternalProjectGroup_Click()
    Dim intRetVal
    Dim strOpenFileName As String
    ' Display a Open dialog box and return a filename.
    On Error Resume Next
    frmMDI.CMDialog1.DialogTitle = "Open Project"
    frmMDI.CMDialog1.FileName = ""
    frmMDI.CMDialog1.Filter = "Project Files (*.etp)"
    frmMDI.CMDialog1.FilterIndex = 1
    frmMDI.CMDialog1.DefaultExt = "etp"
    frmMDI.CMDialog1.Flags = &H4    'fcdlOFNHideReadOnly
    frmMDI.CMDialog1.ShowOpen

    strOpenFileName = frmMDI.CMDialog1.FileName
    ' check if the extension on file is etp
    If Right(strOpenFileName, 3) = "etp" Or Right(strOpenFileName, 3) = "ETP" Then
        'save the projects to file
        Call LoadTreeViewFromFile(strOpenFileName, TVProjects)
    Else
        MsgBox "This is not a E-Type project file!", vbCritical, "Warning"
    End If
End Sub

Private Sub mnuNewGroup_Click()
    Call AddNewGroup(TVProjects)
End Sub

Private Sub mnuPaths_Click(index As Integer)
    If mnuPaths(index).Caption = "History" Then Exit Sub
    'Change the Directory items to equal the new Current Directory
    Dir1.Path = mnuPaths(index).Caption
    cboPath.text = mnuPaths(index).Caption
    Drive1.Drive = Left(cboPath, 3)
    ListView1.ListItems.Clear 'Clear Out Old Items
    'add file and dirnames to the listview
    PopulateListView

End Sub

Private Sub mnuRecentFile_Click(index As Integer)
    ' Call the file open procedure, passing a
    ' reference to the selected file name
    OpenFile (mnuRecentFile(index).Caption)
    ' Update the list of the most recently opened files.
    GetRecentFiles
End Sub

Public Sub DocumentSelChange(rtfText As RichTextBox)
    
    Dim i As Long
    Dim lngCurCol As Long
    Dim lngCurLine As Long
    Dim lngPos As Long
    Dim lngLineCount As Long
    Dim blnEnabled As Boolean
    Dim strStatusText As String
    With rtfText
        ' Line Count
        lngLineCount = SendMessage(.hwnd, EM_GETLINECOUNT, 0&, 0&)
        ' Current Line
        lngCurLine = 1 + .GetLineFromChar(.SelStart)
        ' Current Char
        lngPos = rtfText.SelStart + 1
        ' Column
        i = SendMessage(rtfText.hwnd, EM_LINEINDEX, ByVal lngCurLine - 1, 0&)
        ' column
        lngCurCol = (lngPos) - i
        Select Case .SelLength
        Case Is >= 1
            strStatusText = "Pos: " & lngPos & ":" & .SelLength
        Case 0
            strStatusText = "Pos: " & lngPos
        End Select
        ' Char info
        strStatusText = strStatusText & "  Ln " & lngCurLine & "/" & lngLineCount & "  Col " & lngCurCol & " "
        SetStatusBar strStatusText, "CharNum"
    End With

End Sub

Private Sub mnuRemoveDocument_Click()
    If TVProjects.SelectedItem.Image = "Folder" Then
        MsgBox "Selected object not a document!", vbInformation, "Error"
    Else ' if selected item is a child
        TVProjects.Nodes.Remove TVProjects.SelectedItem.index
    End If

    
End Sub

Private Sub mnuRenameGroup_Click()
    Dim Name As String
    Dim Group As Node
    Dim CurrentName As String
    
    ' Find the current group name.
    ' if the selected item is a parent
    If TVProjects.SelectedItem.Image = "Folder" Then
        CurrentName = TVProjects.SelectedItem.text
    Else ' if selected item is a child
        CurrentName = TVProjects.SelectedItem.Parent.text
    End If
    
    Name = InputBox("New Project Name", , CurrentName)
    If Name = "" Then Exit Sub
    
        ' Find the group for renaming.
        ' if the selected item is a parent
        If TVProjects.SelectedItem.Image = "Folder" Then
            TVProjects.SelectedItem.text = Name
        Else ' if selected item is a child
            TVProjects.SelectedItem.Parent.text = Name
        End If


End Sub

Private Sub mnusaveProjectsGroup_Click()
    Dim strSaveFileName As String
    Dim strDefaultName As String
    ' Assign a default name to the variable.
    strDefaultName = "NewProject.etp"
    
    ' Display a Save As dialog box and return a filename.
    On Error Resume Next
    frmMDI.CMDialog1.DialogTitle = "Save Project"
    frmMDI.CMDialog1.FileName = strDefaultName
    frmMDI.CMDialog1.Filter = "Project Files (*.etp)"
    frmMDI.CMDialog1.FilterIndex = 1
    frmMDI.CMDialog1.DefaultExt = "etp"
    frmMDI.CMDialog1.Flags = &H4    'fcdlOFNHideReadOnly
    frmMDI.CMDialog1.ShowSave
       
    strSaveFileName = frmMDI.CMDialog1.FileName
    'save the projects to file
    Call SaveTreeViewToFile(strSaveFileName, TVProjects)

End Sub


Private Sub mnusaveProjectsGroupNow_Click()
    'save the projects to file
    Call SaveTreeViewToFile((App.Path) & "\" & "projects.etp", TVProjects)
End Sub

Private Sub mnuTabs_Click()
    Me.tabDocuments.Style = tabTabs
End Sub

Private Sub sstDriveFilelist_Click(PreviousTab As Integer)
    ' this sub sets the text on lblTabInfo acording to wich tab the user select
    Dim myIndex As Integer
    myIndex = sstDriveFilelist.Tab
    Select Case myIndex
        Case 0
            lblTabInfo.Caption = "Insert"
            Case 1
            lblTabInfo.Caption = "Projects"
            Case 2
            lblTabInfo.Caption = "Drive/File list"
            Case 3
            lblTabInfo.Caption = "Find Results"
        End Select
End Sub

Private Sub tabDocuments_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' bring the selected document to front
    ' the tab key contains the three letters "key" before the number representing
    ' the document index. These letters must be removed before we can
    ' refere to the document

    Dim dNum As String
    dNum = (Mid(tabDocuments.SelectedItem.key, 4))
    Document(dNum).SetFocus

End Sub

Private Sub tabDocuments_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' bring the selected document to front
    ' the tab key contains the three letters "key" before the number representing
    ' the document index. These letters must be removed before we can
    ' refere to the document

    Dim dNum As String
    dNum = (Mid(tabDocuments.SelectedItem.key, 4))
    Document(dNum).SetFocus

    ' Show popupmenu
    If Button = vbRightButton Then 'do the popup menu
        Button = vbLeftButton
        PopupMenu mnuDocTabs
    End If

End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Select Case Button.key
        Case "New"
            mnuFileNew_Click
        Case "Open"
            Call FileOpenProc
        Case "Save"
            Call frmMDI.ActiveForm.SaveFile
        Case "Print"
            Call frmMDI.ActiveForm.printText
        Case "PrintPreview"
            frmDocPreview.Show vbModal
        Case "Cut"
            Me.ActiveForm.EditCut
        Case "Copy"
            Me.ActiveForm.EditCopy
        Case "Paste"
            Me.ActiveForm.EditPaste
        Case "Undo"
            Call Me.ActiveForm.Undo
'            Me.ActiveForm.EditUndo
        Case "Redo"
            Call Me.ActiveForm.Redo
'            Me.ActiveForm.EditRedo
        Case "Find"
            Call frmMDI.ActiveForm.SearchFind
        Case "FindN"
            ' Assign a value to the public variable.
            gFindDirection = 1
            Call frmMDI.ActiveForm.SearchFindNext
        Case "FindP"
            ' Assign a value to the public variable.
            gFindDirection = 0
            Call frmMDI.ActiveForm.SearchFindPrev
        Case "Replace"
            Call frmMDI.ActiveForm.SearchReplace
        Case "NextWindow"
            Call SelNextWin
        Case "PrevWindow"
            Call SelPrevWin
        Case "Cascade"
            frmMDI.Arrange vbCascade
        Case "TileH"
            frmMDI.Arrange vbTileHorizontal
        Case "TileV"
            frmMDI.Arrange vbTileVertical
        Case "Align Right"
            ActiveForm.Text1.SelAlignment = rtfRight
        Case "SelVindow"
            ' add/remove the new window to the toolbar windowlist button
            Call RebuildWinList
        Case "Properties"
            frmFileInfo.Show
    End Select

End Sub

Private Sub tbToolBar_ButtonMenuClick(ByVal Button As MSComctlLib.ButtonMenu)
    Dim i As Long

    For i = 1 To Forms.Count - 1
        If Button.text = Forms(i).Caption Then
            Forms(i).SetFocus
        End If
    Next
'    Dim i As Integer        ' Counter variable
'
'    ' Cycle through the document array.
'    ' Return true if there is at least one open document.
'    For i = 1 To UBound(Document)
'        If Button.Text = Document(i).Caption Then
'            Document(i).SetFocus
'        End If
'    Next
End Sub

Public Sub DrivListStart()

    ' This function deals with the file/drivelist on picSideBar
    
    ' Create an object variable for the ColumnHeader object.
    ' Add ColumnHeaders.  The width of the columns is the width
    ' of the control divided by the number of ColumnHeader objects.
    ListView1.ColumnHeaders.Add , , "Name", 2000
    Set clmX = ListView1.ColumnHeaders.Add(, , "Size", ListView1.Width / 3, lvwColumnRight)
    Set clmX = ListView1.ColumnHeaders.Add(, , "Date", 1500)
    
    ' To use ImageList controls with the ListView control, you must
    ' associate a particular ImageList control with the Icons and
    ' Icons were previously Added to list
    
    ' SmallIcons properties.
    ListView1.Icons = imlDriveFileList2
    ListView1.SmallIcons = imlDriveFileList2

    'Start Off With Current Drive and directory
    ChDrive Drive1.Drive
    Dir1.Path = CurDir
    
    'Adding items to the combobox
    cboFiletypes.AddItem "Ascii Text (*.txt)", 0
    cboFiletypes.AddItem "Rich Text (*.rtf)", 1
    cboFiletypes.AddItem "Bat Files (*.bat)", 2
    cboFiletypes.AddItem "Ini Files (*.ini)", 3
    cboFiletypes.AddItem "Sys Files (*.sys)", 4
    cboFiletypes.AddItem "All Files (*.*)", 5

End Sub

Public Sub PopulateListView()
    ' This function populates the listview on picSideBar
    
    'If we are in a subdirectory then do the following
    If Right(Dir1.Path, 1) <> "\" Then
        CurrentDir = Dir1.Path & "\"
        dname = ".."
        Set itmX = ListView1.ListItems.Add(, , dname)
        itmX.SubItems(1) = ""
        itmX.Icon = 5          ' Set an icon from imlDriveFileList.
        itmX.SmallIcon = 5      ' Set an icon from ImageList2.
        itmX.SubItems(2) = ""
    Else
        'If not in a subdirectory then do the following
        CurrentDir = Dir1.Path
    End If

    'Get the Directory Names first
    For counter = 0 To Dir1.ListCount - 1
        dname = Dir1.list(counter)
        For Counter2 = Len(dname) To 1 Step -1
            If Mid$(dname, Counter2, 1) = "\" Then
                TempDname = Right(dname, Len(dname) - Counter2)
                Exit For
            End If
        Next Counter2
        Set itmX = ListView1.ListItems.Add(, , TempDname)
        itmX.SubItems(1) = ""
        itmX.Icon = 3          ' Set an icon from imlDriveFileList.
        itmX.SmallIcon = 3      ' Set an icon from ImageList2.
        itmX.SubItems(2) = FileDateTime(dname)
    Next counter
    
    'Get the FileNames next
    For counter = 0 To File1.ListCount - 1
        Fname = File1.list(counter)
        Set itmX = ListView1.ListItems.Add(, , Fname)
        itmX.SubItems(1) = CStr(FileLen(CurrentDir & Fname))
        itmX.Icon = 4           ' Set an icon from imlDriveFileList.
        itmX.SmallIcon = 4      ' Set an icon from ImageList2.
        itmX.SubItems(2) = FileDateTime(CurrentDir & Fname)
    Next counter
    
    '**************************
    Dim duplicate As Boolean
        
    duplicate = frmFind.CheckDup(Dir1.Path, cboPath)
    If duplicate = False Then
        cboPath.AddItem Dir1.Path
        'Selecting the first item
        cboPath.ListIndex = Str(cboPath.ListCount) - 1
        cboPath.ToolTipText = Dir1.Path
    End If
    updateMnuPaths (Dir1.Path)
End Sub



Public Sub PrepFindResults()
    ' This sub deals with the find results on picSideBar
    
    ' Create an object variable for the ColumnHeader object.
    ' Add ColumnHeaders.  The width of the columns is the width
    ' of the control divided by the number of ColumnHeader objects.
    Dim clmX As ColumnHeader
    
    lvFindResults.ColumnHeaders.Add , , "Name", 2000
    Set clmX = lvFindResults.ColumnHeaders.Add(, , "Size", lvFindResults.Width / 3)
    Set clmX = lvFindResults.ColumnHeaders.Add(, , "Location", 8500)
    
    ' To use ImageList controls with the ListView control, you must
    ' associate a particular ImageList control with the Icons and
    ' Icons were previously Added to list
    
    ' SmallIcons properties.
    lvFindResults.Icons = imlDriveFileList2
    lvFindResults.SmallIcons = imlDriveFileList2


End Sub

Private Sub tbToolBar_Change()
    ' Save toolbarsettings to registry
    'tbToolBar.SaveToolbar App.Title, "Toolbar", "Toolbar1"
    tbToolBar.SaveToolbar App.Title, "Toolbar", "Toolbar1"


End Sub

Private Sub TVProjects_DblClick()
    ' if selected item is a parrent
    If TVProjects.SelectedItem.Image = "Folder" Then
    ' nothing
    Else ' if selected item is a child
        ' Call the file open procedure, passing a
        ' reference to the selected file name
        OpenFile TVProjects.SelectedItem.text
        ' Update the list of recently opened files in the File menu control array.
        UpdateFileMenu TVProjects.SelectedItem
        If SBarMain.Visible = True Then
            'Get the filestats from the current active file
            GetFileStats
        End If
    End If
End Sub

Private Sub TVProjects_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' first check if there is any nodes at all, then
    ' check if the selected node is folder or document.
    ' we should not let the user delete a folder if a
    ' document is selected.

    If Button = vbRightButton Then 'do the popup menu
        ' if there is no nodes in treeview
        If TVProjects.Nodes.Count = 0 Then
            mnuAddDocument.Enabled = False
            mnuAddCurrentDocument.Enabled = False
            mnuRemoveDocument.Enabled = False
            mnuNewGroup.Enabled = True
            mnuDeleteGroup.Enabled = False
            mnuRenameGroup.Enabled = False
            mnusaveProjectsGroup.Enabled = False
            mnuLoadExsternalProjectGroup.Enabled = True
            ' Show popupmenu
            PopupMenu mnuProjects
            Exit Sub
        End If
        ' if selected item is a parrent
        If TVProjects.SelectedItem.Image = "Folder" Then
            mnuAddDocument.Enabled = True
            mnuAddCurrentDocument.Enabled = True
            mnuRemoveDocument.Enabled = False
            mnuNewGroup.Enabled = True
            mnuDeleteGroup.Enabled = True
            mnuRenameGroup.Enabled = True
            mnusaveProjectsGroup.Enabled = True
            mnuLoadExsternalProjectGroup.Enabled = True
        Else ' if selected item is a child
            mnuAddDocument.Enabled = True
            mnuAddCurrentDocument.Enabled = True
            mnuRemoveDocument.Enabled = True
            mnuNewGroup.Enabled = True
            mnuDeleteGroup.Enabled = False
            mnuRenameGroup.Enabled = True
            mnusaveProjectsGroup.Enabled = True
            mnuLoadExsternalProjectGroup.Enabled = True
        End If
        ' Show popupmenu
        PopupMenu mnuProjects
    End If

End Sub

Public Sub SelNextWin()
    ' purpose: Select the next window in the document array.
    ' Allso, checks if the next window in the array is not deleted.
    ' If it is, check if there are a next window in the array. If not
    ' start the search for next window from document(1), again check to
    ' see if it exists.
    
    Dim Current As Integer
    Dim i As Integer
    
    Current = Me.ActiveForm.Tag
    
    ' Cycle through the document array. If one of the
    ' documents has been deleted, step to next.
    For i = Current + 1 To UBound(Document)
        If FState(i).Deleted Then
            While FState(i).Deleted
                i = i + 1
                If i > UBound(Document) Then
                    Exit For ' jump out of the loop and start the search at document(1)
                End If
                
            Wend
            Document(i).SetFocus
            Exit Sub
        Else
            Document(i).SetFocus
            Exit Sub
        End If
    Next
        
    i = LBound(Document)
    While FState(i).Deleted = True
        i = i + 1
    Wend
    Document(i).SetFocus
    Exit Sub
End Sub

Public Sub SelPrevWin()
    ' purpose: Select the previous window in the document array.
    ' Allso, checks if the previous window in the array is not deleted.
    ' If it is, check if there are a previous window in the array. If not
    ' start the search for next window from the top of the array,
    ' again check to see if it exists.
    
    Dim Current As Integer
    Dim i As Integer
    
    Current = Me.ActiveForm.Tag
    
    ' Cycle through the document array. If one of the
    ' documents has been deleted, step to next.
    For i = Current - 1 To LBound(Document) Step -1
        If FState(i).Deleted Then
            While FState(i).Deleted
                i = i - 1
                If i = (LBound(Document) - 1) Then
                    Exit For ' jump out of the loop and start the search at the top of the document array
                End If
                
            Wend
            Document(i).SetFocus
            Exit Sub
        Else
            Document(i).SetFocus
            Exit Sub
        End If
    Next
        
    i = UBound(Document)
    While FState(i).Deleted = True
        i = i - 1
    Wend
    Document(i).SetFocus
    Exit Sub

'    Dim Current As Integer
'    Dim i As Integer
'    Current = Me.ActiveForm.Tag
'    For i = Current - 1 To 1 Step -1
'        Document(i).SetFocus
'            Exit Sub
'    Next
'    For i = UBound(Document) To Current + 1 Step -1
'        Document(i).SetFocus
'        Exit Sub
'    Next
End Sub

Public Sub resizeME()
    ' Scaling all the controls in picSideBar to match the
    ' main forms height
    Dim freespace As Integer
    
    If tbToolBar.Visible = True Then freespace = tbToolBar.Height
    If SBarMain.Visible = True Then freespace = freespace + SBarMain.Height
    If tbaDocuments.Visible = True Then freespace = freespace + tbaDocuments.Height
    If picSideBar.Visible = True Then   ' no need to resize if not invisible
        If Me.Height > 2600 Then        ' just as long as formheigh is more than 2600
            picContainer.Height = Me.Height
            sstDriveFilelist.Height = Me.Height - (freespace + 1000)
            cboFiletypes.Top = sstDriveFilelist.Height - 380
            ListView1.Height = sstDriveFilelist.Height - 1190
            lvFindResults.Height = sstDriveFilelist.Height - 200
            TVProjects.Height = sstDriveFilelist.Height - 200
            lstLibraries.Height = sstDriveFilelist.Height - 500
        End If
    End If
    ' scaling tabDocuments
    tabDocuments.Width = Me.Width - 100

End Sub

'************************************************************
'If any windows are open or closed this routine is called to
'set the status of certain items
'************************************************************
Public Sub ChildStatusChanged()
    Dim lbStatus As Boolean
    
    RebuildWinList
    
    lbStatus = (Forms.Count > 2)
    tbToolBar.Buttons("NextWindow").Enabled = lbStatus
    tbToolBar.Buttons("PrevWindow").Enabled = lbStatus
    tbToolBar.Buttons("TileH").Enabled = lbStatus
    tbToolBar.Buttons("TileV").Enabled = lbStatus
    tbToolBar.Buttons("Cascade").Enabled = lbStatus
    tbToolBar.Buttons("SelWindow").Enabled = lbStatus
       
    lbStatus = (Forms.Count > 1)
End Sub


Public Sub updateMnuPaths(PathName As String)
        Dim i As Integer    ' Counter variable.
        Dim x As Integer    ' Counter variable
        
        pathMenuCount = mnuPaths.UBound + 1
        ' Check if the pathname already exists in the menu control array.
        For i = mnuPaths.LBound To mnuPaths.UBound
            ' Check if the pathname already exists in the menu control array.
            If mnuPaths(i).Caption = PathName Then
                Exit Sub
            End If
        Next i
            If pathMenuCount = 0 Then   ' no entries in menu
                mnuPaths(0).Visible = True
                mnuPaths(0).Caption = PathName
            Else
                Load mnuPaths(pathMenuCount)
                mnuPaths(pathMenuCount).Visible = True
                mnuPaths(pathMenuCount).Caption = PathName
            End If

'        Next i
End Sub

Public Sub StartSideBar()
    
    ' populate the filelist with files
    DrivListStart
    
    ' populate the projecttreeview with projects read from the file projects.etp
    Call LoadTreeViewFromFile((App.Path) & "\" & "projects.etp", TVProjects)
    
    ' prepare the lvFindResults
    PrepFindResults
    
    ' select container 3 on sstDriveFileList
    frmMDI.sstDriveFilelist.Tab = 2
    
    Me.resizeME
     
    ' show the items on the sidebar. Hiding them until now makes
    ' the loading of the main form nicer to look at.
    sstDriveFilelist.Visible = True
    
    ' Selecting the first item in cbofiletypes. This will also
    ' populate the listview
    cboFiletypes.ListIndex = 5
    
End Sub

Public Sub dupeWindow()
    ' Save activeform file as e-type.tmp in temp directory
    frmMDI.ActiveForm.Text1.SaveFile GetTmpPath & "\r-type.tmp", rtfText
    ' Call the new form procedure
    FileNew
    ' Load e-type.tmp from temp directory into new window
    frmMDI.ActiveForm.Text1.LoadFile GetTmpPath & "\r-type.tmp", rtfText
    ' Delete tempfile
    Kill GetTmpPath & "\r-type.tmp"

End Sub
