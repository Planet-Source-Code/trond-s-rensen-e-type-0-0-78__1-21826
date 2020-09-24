VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{307C5043-76B3-11CE-BF00-0080AD0EF894}#1.0#0"; "MSGHOO32.OCX"
Begin VB.Form frmDriveFilelist 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Drive/Filelist"
   ClientHeight    =   6525
   ClientLeft      =   4155
   ClientTop       =   1875
   ClientWidth     =   2445
   Icon            =   "frmDriveFilelist.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   2445
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3120
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDriveFilelist.frx":0C42
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDriveFilelist.frx":0D56
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDriveFilelist.frx":0E6A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.FileListBox File1 
      Height          =   675
      Left            =   2640
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1560
      Width           =   2655
   End
   Begin VB.DirListBox Dir1 
      Height          =   765
      Left            =   2640
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   720
      Width           =   2655
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6495
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   11456
      _Version        =   393216
      Style           =   1
      Tab             =   2
      TabHeight       =   520
      WordWrap        =   0   'False
      TabCaption(0)   =   " "
      TabPicture(0)   =   "frmDriveFilelist.frx":0F7E
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmbLibraries"
      Tab(0).Control(1)=   "lstLibraries"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   " "
      TabPicture(1)   =   "frmDriveFilelist.frx":1310
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   " "
      TabPicture(2)   =   "frmDriveFilelist.frx":159E
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "cboFiletypes"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "ListView1"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Drive1"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   30
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   360
         Width           =   2355
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   5370
         Left            =   30
         TabIndex        =   6
         Top             =   705
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   9472
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
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
      Begin VB.ListBox lstLibraries 
         Appearance      =   0  'Flat
         Height          =   5685
         ItemData        =   "frmDriveFilelist.frx":1900
         Left            =   -74950
         List            =   "frmDriveFilelist.frx":1907
         TabIndex        =   5
         Top             =   720
         Width           =   2330
      End
      Begin VB.ComboBox cmbLibraries 
         Height          =   315
         Left            =   -74950
         TabIndex        =   4
         Text            =   "Combo2"
         Top             =   360
         Width           =   2330
      End
      Begin VB.ComboBox cboFiletypes 
         Height          =   315
         ItemData        =   "frmDriveFilelist.frx":1910
         Left            =   30
         List            =   "frmDriveFilelist.frx":1912
         TabIndex        =   3
         Top             =   6120
         Width           =   2350
      End
   End
   Begin MsghookLib.Msghook MsgHook 
      Left            =   3240
      Top             =   5400
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   2760
      TabIndex        =   8
      Top             =   2640
      Width           =   2055
   End
End
Attribute VB_Name = "frmDriveFilelist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*
' Chetan Sarva
' csarva@ic.sunysb.edu
'
' This code is a variation of the code posted by
' Steve of http://www.vbtutor.com/ on Planet
' Source Code. This code requires the
' MsgHook 32 OCX available here:
' http://www.mvps.org/vb/code/msghook.zip
' or possibly from somewhere on the Mabry
' site as well.
'
' If you use this code in any way, shape or
' form, please include all the above lines
' somewhere at the top of your application.
' It is only fair to give credit where credit is due.

' ####################################
' Declared Functions, Constants, and Types
' ####################################

'API Declares
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As Rect, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As Rect) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As Any) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As Rect, ByVal lLeft As Long, ByVal lTop As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal Y As Long) As Long
Private Declare Function SetROP2 Lib "gdi32" (ByVal hDC As Long, ByVal nDrawMode As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal H%, ByVal hb%, ByVal x%, ByVal Y%, ByVal Cx%, ByVal Cy%, ByVal f%) As Integer
Private Declare Function SetWindowWord Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal wNewWord As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

' Constants

Private Const WM_NCACTIVATE = &H86
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const HWND_NOTOPMOST = -2
Private Const HWND_TOPMOST = -1
Private Const WM_SYSCOMMAND = &H112
Private Const VK_LBUTTON = &H1
Private Const PS_SOLID = 0
Private Const R2_NOTXORPEN = 10
Private Const BLACK_PEN = 7

' Types

Private Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    x As Long
    Y As Long
End Type

' #############
' User Variables
' #############

'Public variables used elsewhere to set values
' for this form's position and size.
Dim lFloatingWidth As Long
Dim lFloatingHeight As Long
Dim lFloatingLeft As Long
Dim lFloatingTop As Long

Dim bMoving As Boolean

'Private variables used to track moving/sizing etc.
Public bDocked As Boolean
Public lDockedWidth As Long
Public lDockedHeight As Long

Dim fLeft As Long ' Hold the form's x coordinate for the Form_Moved event
Dim fWidth As Long ' Hold the form width for the Form_Moved event
Dim fHeight As Long ' Hold the form height for the Form_Moved event

Dim dropZone As Integer ' Size of the drop area

Dim TitleBarHeight As Integer ' Hold the height of the titlebar of our mdi form
Dim dockParent As Long ' Hold the hwnd of the mdi parent window
'*

'Variables to workaround the fact that there is no ItemDblClick Event

Dim xpos As Long
Dim ypos As Long

Dim NameOfFile As String
Dim clmX As ColumnHeader
Dim Item As ListItem
Dim itmX As ListItem
Dim Counter As Long
Dim Counter2 As Integer
Dim dname As String
Dim Fname As String
Dim TempDname As String
Dim CurrentDir As String



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

    'add file and dirnames to the listview
    PopulateListView

End Sub


Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()

    ChDrive Drive1.Drive
    Dir1.Path = Drive1.Drive
    
    ListView1.ListItems.Clear 'Clear Out Old Items

    'add file and dirnames to the listview
    PopulateListView

End Sub

Private Sub Form_Load()

    '*
    Dim mdiRect As Rect
    
    frmDFLLoaded = True
    
    ' Initialize the drop area to a default value
    '*dropZone = 1500
    dropZone = frmDriveFilelist.Width
    ' Get the screen based coordinates of our MDI Form
    GetWindowRect frmMDI.hWnd, mdiRect
    ' Calculate the height of the titlebar and store it in a
    ' variable. We'll need this for later when a form is docked.
    TitleBarHeight = (mdiRect.Bottom - mdiRect.Top - frmMDI.ScaleHeight \ Screen.TwipsPerPixelY) - 8
    
    ' The hWnd of the window that is the actual parent
    ' of our form is different from frmMDI.hWnd
    ' Get it, and store it in a variable for later use
    dockParent = GetParent(Me.hWnd)
    
    ' Initialize the positions/sizes of this form
    lFloatingLeft = Me.Left
    lFloatingTop = Me.Top
    lFloatingWidth = Me.Width
    lFloatingHeight = Me.Height
    
    ' Subclass this form
    With MsgHook
        .HwndHook = Me.hWnd
        .Message(WM_SYSCOMMAND) = True
    End With
    '*
    
    
    ' Create an object variable for the ColumnHeader object.
    ' Add ColumnHeaders.  The width of the columns is the width
    ' of the control divided by the number of ColumnHeader objects.
    ListView1.ColumnHeaders.Add , , "Name", 2000
    Set clmX = ListView1.ColumnHeaders.Add(, , "Size", ListView1.Width / 3)
    Set clmX = ListView1.ColumnHeaders.Add(, , "Date", 1500)

    ' To use ImageList controls with the ListView control, you must
    ' associate a particular ImageList control with the Icons and
    ' Icons were previously Added to list
    
    ' SmallIcons properties.
    ListView1.Icons = ImageList1
    ListView1.SmallIcons = ImageList1


    'Start Off With Current Drive and directory
    ChDrive Drive1.Drive
    Dir1.Path = CurDir
    
    'Add BackSlash if Necessary
    'If Right(CurrentDir, 1) <> "\" Then CurrentDir = CurrentDir & "\"
    'Dir1.Path = CurrentDir
    'Dir1.Path = Drive1.Drive
    
    'NameOfFile = Dir$(CurrentDir & "*.*", vbDirectory)
    
    'add file and dirnames to the listview
    PopulateListView

    'Adding items to the combobox
    cboFiletypes.AddItem "Ascii Text (*.txt)", 0
    cboFiletypes.AddItem "Rich Text (*.rtf)", 1
    cboFiletypes.AddItem "Bat Files (*.bat)", 2
    cboFiletypes.AddItem "Ini Files (*.ini)", 3
    cboFiletypes.AddItem "Sys Files (*.sys)", 4
    cboFiletypes.AddItem "All Files (*.*)", 5
    'Selecting the first item
    cboFiletypes.ListIndex = 0
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '*
    If bDocked Then
        bDocked = False
        SetParent Me.hWnd, frmMDI.hWnd
    End If

End Sub

Private Sub Form_Resize()
'*
    ' Update the stored Values
    If Me.WindowState <> vbMinimized Then StoreFormDimensions

End Sub

Private Sub Form_Unload(Cancel As Integer)
    '*
    frmDFLLoaded = False
    
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    
    ListView1.SortKey = ColumnHeader.Index - 1

End Sub

Private Sub ListView1_DblClick()
        
    If ListView1.HitTest(xpos, ypos) Is Nothing Then
        Exit Sub
    Else
        Set Item = ListView1.HitTest(xpos, ypos)
    End If
    'Label1.Caption = Item
    ' Call the file open procedure, passing a
    ' reference to the selected file name
    OpenFile Item
    ' Update the list of recently opened files in the File menu control array.
    GetRecentFiles
    'Get the filestats from the current active file
    GetFileStats
'    frmNotePad.Text1.LoadFile Item
    'Set Item = ListView1.SelectedItem
    
    'If you Click on a filename just exit this subroutine
    If Right(Dir1.Path, 1) <> "\" Then
        CurrentDir = Dir1.Path & "\"
    Else
        CurrentDir = Dir1.Path
    End If
    
    If (GetAttr(CurrentDir & Item) And vbDirectory) <= 0 Then Exit Sub
    ListView1.ListItems.Clear 'Clear Out Old Items
    
    'Change to selected Directory - Let Visual Basic do the work
    ChDir Item
    
    'Change the Directory List Box to equal the new Current Directory
    Dir1.Path = CurDir
    
    'add file and dirnames to the listview
    PopulateListView

End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

    xpos = x
    ypos = Y

End Sub

Public Sub PopulateListView()
    
    'If we are in a subdirectory then do the following
    If Right(Dir1.Path, 1) <> "\" Then
        CurrentDir = Dir1.Path & "\"
        dname = ".."
        Set itmX = ListView1.ListItems.Add(, , dname)
        itmX.SubItems(1) = ""
        itmX.Icon = 3           ' Set an icon from ImageList1.
        itmX.SmallIcon = 3      ' Set an icon from ImageList2.
        itmX.SubItems(2) = ""
    Else
        'If not in a subdirectory then do the following
        CurrentDir = Dir1.Path
    End If

    'Get the Directory Names first
    For Counter = 0 To Dir1.ListCount - 1
        dname = Dir1.List(Counter)
        For Counter2 = Len(dname) To 1 Step -1
            If Mid$(dname, Counter2, 1) = "\" Then
                TempDname = Right(dname, Len(dname) - Counter2)
                Exit For
            End If
        Next Counter2
        Set itmX = ListView1.ListItems.Add(, , TempDname)
        itmX.SubItems(1) = ""
        itmX.Icon = 1           ' Set an icon from ImageList1.
        itmX.SmallIcon = 1      ' Set an icon from ImageList2.
        itmX.SubItems(2) = FileDateTime(dname)
    Next Counter
    
    'Get the FileNames next
    For Counter = 0 To File1.ListCount - 1
        Fname = File1.List(Counter)
        Set itmX = ListView1.ListItems.Add(, , Fname)
        itmX.SubItems(1) = CStr(FileLen(CurrentDir & Fname))
        itmX.Icon = 2           ' Set an icon from ImageList1.
        itmX.SmallIcon = 2      ' Set an icon from ImageList2.
        itmX.SubItems(2) = FileDateTime(CurrentDir & Fname)
    Next Counter

End Sub

'*
Private Sub MsgHook_Message(ByVal msg As Long, ByVal wp As Long, ByVal lp As Long, result As Long)
    
    'Debug.Print GetWinMsgStr(msg)
    Debug.Print wp
    
    Select Case msg
        Case WM_SYSCOMMAND
            
            ' User is dragging the form. Simulate it with API
            If (wp = 61458) Then
                
                ' Local variables
                Dim pt As POINTAPI ' Current mouse location
                Dim ptPrev As POINTAPI ' Previous mouse location
                Dim mdiRect As Rect ' MDI Rectangle
                Dim objRect As Rect ' Form Rectangle
                Dim DragRect As Rect ' New Rectangle
                Dim lBorderWidth As Long ' Width of the border of the form (normally 3)
                Dim lObjWidth As Long ' Width of our form
                Dim lObjHeight As Long ' Height of our form
                Dim lXOffset As Long ' X Offset from mouse location
                Dim lYOffset As Long ' Y Offset from mouse location
                Dim bMoved As Boolean ' Did the form move?
                
                ' Get the rectangles for our MDI and our Form
                GetWindowRect frmMDI.hWnd, mdiRect
                GetWindowRect Me.hWnd, objRect
                
                ' Determine the height and width of our form
                lObjWidth = objRect.Right - objRect.Left
                lObjHeight = objRect.Bottom - objRect.Top
                
                ' Get the location of our cursor and...
                GetCursorPos pt
                ' ... store it
                ptPrev.x = pt.x
                ptPrev.Y = pt.Y
                
                ' Determine offsets
                lXOffset = pt.x - objRect.Left
                lYOffset = pt.Y - objRect.Top
                
                ' Create inital rectangle for drawing form's "edges"
                With DragRect
                    .Left = pt.x - lXOffset
                    .Top = pt.Y - lYOffset
                    .Right = .Left + lObjWidth
                    .Bottom = .Top + lObjHeight
                End With
                
                ' Width of the form's border - used for showing
                ' the user that form is being dragged
                lBorderWidth = 3
                ' Draw the rectangle on the screen
                DrawDragRectangle DragRect.Left, DragRect.Top, DragRect.Right, DragRect.Bottom, lBorderWidth
                
                ' Check if the form is being moved
                Do While GetKeyState(VK_LBUTTON) < 0
                    GetCursorPos pt
                    
                    If (pt.x <> ptPrev.x Or pt.Y <> ptPrev.Y) Then
                       
                       If pt.x < mdiRect.Left + 6 Then
                            pt.x = mdiRect.Left + 6
                        ElseIf pt.x > mdiRect.Right - 6 Then
                            pt.x = mdiRect.Right - 6
                        End If
                                            
                        If pt.Y < mdiRect.Top + TitleBarHeight + 3 Then
                            pt.Y = mdiRect.Top + TitleBarHeight + 3
                        ElseIf pt.Y > mdiRect.Bottom - 6 Then
                            pt.Y = mdiRect.Bottom - 6
                        End If
                        
                        SetCursorPos pt.x, pt.Y
                        
                        ptPrev.x = pt.x
                        ptPrev.Y = pt.Y
                        
                        ' Erase the previous drag rectangle, if it exists, by drawing on top of it
                        DrawDragRectangle DragRect.Left, DragRect.Top, DragRect.Right, DragRect.Bottom, lBorderWidth
                        
                        ' Fire the Form_Moved event
                        Call Form_Moved(pt.x, pt.Y)
                        
                        ' Adjust the height/width
                        With DragRect
                            If fLeft > 0 Then .Left = fLeft Else .Left = pt.x - lXOffset
                            .Top = pt.Y - lYOffset
                            .Right = .Left + fWidth
                            .Bottom = .Top + fHeight
                        End With
                        
                        ' Draw the rectagle again at it's new position and dimensions
                        DrawDragRectangle DragRect.Left, DragRect.Top, DragRect.Right, DragRect.Bottom, lBorderWidth
                        bMoved = True
                    End If ' (pt.X <> ptPrev.X Or pt.Y <> ptPrev.Y)
                    
                    DoEvents
                Loop ' While GetKeyState(VK_LBUTTON) < 0
                
                ' Erase the previous drag rectangle if any
                DrawDragRectangle DragRect.Left, DragRect.Top, DragRect.Right, DragRect.Bottom, lBorderWidth
                
                ' User let go of the left mouse button.
                ' If it is in a new location, fire Form_Dropped
                If (bMoved) Then
                    
                    Call Form_Dropped(pt.x, pt.Y)
                    
                    ' If the form isn't docked, move it to it's new location
                    If Not bDocked Then MoveWindow Me.hWnd, DragRect.Left - mdiRect.Left - 6, DragRect.Top - mdiRect.Top + 6 - (mdiRect.Bottom - mdiRect.Top - frmMDI.ScaleHeight \ Screen.TwipsPerPixelY), DragRect.Right - DragRect.Left, DragRect.Bottom - DragRect.Top, True
                
                End If ' (bMoved)
            
            ' User clicked the close button. First we have to
            ' set the parent back to the MDI or it will lock up.
            ElseIf wp = 61536 Then
                Call SetParent(Me.hWnd, dockParent)
                Me.Visible = False
                frmMDI!Picture1.Visible = False
                Unload Me
                
            ' If we got another wparam, return control to the window
            Else
                result = MsgHook.InvokeWindowProc(msg, wp, lp)
                
            End If ' (wp = 61458)
            
        ' On all other messages, return control the window
        Case Else
            result = MsgHook.InvokeWindowProc(msg, wp, lp)
                
    End Select

End Sub
'*
Sub Calc_Bottom(ptx As Long)

    fLeft = 0
    fWidth = frmMDI.ScaleWidth / Screen.TwipsPerPixelX
    fHeight = dropZone / Screen.TwipsPerPixelY
            
End Sub
'*
Sub Calc_Default()

    fLeft = 0
    fWidth = lFloatingWidth / Screen.TwipsPerPixelX
    fHeight = lFloatingHeight / Screen.TwipsPerPixelY
            
End Sub
'*
Sub Calc_Left(ptx As Long)

    fLeft = ptx - (dropZone / Screen.TwipsPerPixelX) / 2
    fWidth = dropZone / Screen.TwipsPerPixelX
    fHeight = frmMDI.ScaleHeight / Screen.TwipsPerPixelY
            
End Sub
'*
Sub Calc_Right(ptx As Long)

    fLeft = ptx - (dropZone / Screen.TwipsPerPixelX) / 2
    fWidth = dropZone / Screen.TwipsPerPixelX
    fHeight = frmMDI.ScaleHeight / Screen.TwipsPerPixelY
            
End Sub
'*
Sub Calc_Top(ptx As Long)

    fLeft = 0
    fWidth = frmMDI.ScaleWidth / Screen.TwipsPerPixelX
    fHeight = dropZone / Screen.TwipsPerPixelY
            
End Sub
'*
Sub Dock_Bottom()

    frmMDI.Picture1.Align = 2
    frmMDI.Picture1.Height = dropZone
    lDockedWidth = frmMDI.Picture1.ScaleWidth + (8 * Screen.TwipsPerPixelX)
    lDockedHeight = frmMDI.Picture1.ScaleHeight + (8 * Screen.TwipsPerPixelY)
    bDocked = True
    Call SetParent(Me.hWnd, frmMDI!Picture1.hWnd)
    Me.Move -4 * Screen.TwipsPerPixelX, -2 * Screen.TwipsPerPixelY, lDockedWidth - 35, lDockedHeight - 35
    frmMDI!Picture1.Visible = True
    
End Sub
'*
Sub Dock_Left()

    frmMDI.Picture1.Align = 3
    frmMDI.Picture1.Width = dropZone
    lDockedWidth = frmMDI.Picture1.ScaleWidth + (8 * Screen.TwipsPerPixelX)
    lDockedHeight = frmMDI.Picture1.ScaleHeight + (8 * Screen.TwipsPerPixelY)
    bDocked = True
    Call SetParent(Me.hWnd, frmMDI!Picture1.hWnd)
    Me.Move -4 * Screen.TwipsPerPixelX, -2 * Screen.TwipsPerPixelY, lDockedWidth - 35, lDockedHeight - 35
    frmMDI!Picture1.Visible = True
    
End Sub
'*
Sub Dock_Right()

    frmMDI.Picture1.Align = 4
    frmMDI.Picture1.Width = dropZone
    lDockedWidth = frmMDI.Picture1.ScaleWidth + (8 * Screen.TwipsPerPixelX)
    lDockedHeight = frmMDI.Picture1.ScaleHeight + (8 * Screen.TwipsPerPixelY)
    bDocked = True
    Call SetParent(Me.hWnd, frmMDI!Picture1.hWnd)
    Me.Move -4 * Screen.TwipsPerPixelX, -2 * Screen.TwipsPerPixelY, lDockedWidth - 35, lDockedHeight - 35
    frmMDI!Picture1.Visible = True
    
End Sub
'*
Sub Dock_Top()

    frmMDI.Picture1.Align = 1
    frmMDI.Picture1.Height = dropZone
    lDockedWidth = frmMDI.Picture1.ScaleWidth + (8 * Screen.TwipsPerPixelX)
    lDockedHeight = frmMDI.Picture1.ScaleHeight + (8 * Screen.TwipsPerPixelY)
    bDocked = True
    Call SetParent(Me.hWnd, frmMDI!Picture1.hWnd)
    Me.Move -4 * Screen.TwipsPerPixelX, -2 * Screen.TwipsPerPixelY, lDockedWidth - 35, lDockedHeight - 35
    frmMDI!Picture1.Visible = True
    
End Sub
'*
Private Sub Form_Dropped(ptx As Long, pty As Long)
    
    Dim formRect As Rect
    Dim mdiRect As Rect
    Dim picRect As Rect
    Dim leftDock As Rect
    Dim rightDock As Rect
    Dim topDock As Rect
    Dim botDock As Rect

    'If over Picture1 on frmMDI which we are using as a Dock, set parent
    'of this form to Picture1, and position it at -4,-4 pixels, otherwise
    'set this Form's parent to the desktop and postion it at Left,Top
    'We dont need to size the form, as the DragForm control will have done
    'this for us.
    'For the purposes of this example, we only dock if the top left corner
    'of this form is within the area bounded by Picture1
    
    ' Get the screen based coordinates of our MDI Form
    GetWindowRect frmMDI.hWnd, mdiRect
    GetWindowRect frmMDI.Picture1.hWnd, picRect
    
    ' Set up the drop zone regions. These will be used for
    ' check to see if the form is to be docked or not.
    
    With leftDock
        .Left = mdiRect.Left + 4
        .Top = mdiRect.Top + TitleBarHeight '(mdiRect.Bottom - mdiRect.Top - frmMDI.ScaleHeight \ Screen.TwipsPerPixelY) - 8
        .Right = .Left + dropZone \ Screen.TwipsPerPixelX
        .Bottom = .Top + frmMDI.ScaleHeight \ Screen.TwipsPerPixelY + 4
    End With
    
    With rightDock
        .Left = mdiRect.Right - dropZone \ Screen.TwipsPerPixelX - 4
        .Top = mdiRect.Top + TitleBarHeight '(mdiRect.Bottom - mdiRect.Top - frmMDI.ScaleHeight \ Screen.TwipsPerPixelY) - 8
        .Right = mdiRect.Right - 4
        .Bottom = .Top + frmMDI.ScaleHeight \ Screen.TwipsPerPixelY + 4
    End With
    
    With topDock
        .Left = mdiRect.Left + dropZone \ Screen.TwipsPerPixelX + 4 + 1
        .Top = mdiRect.Top + TitleBarHeight '(mdiRect.Bottom - mdiRect.Top - frmMDI.ScaleHeight \ Screen.TwipsPerPixelY) - 8
        .Right = rightDock.Right - dropZone \ Screen.TwipsPerPixelX
        .Bottom = .Top + dropZone \ Screen.TwipsPerPixelX
    End With
    
    With botDock
        .Left = mdiRect.Left + dropZone \ Screen.TwipsPerPixelX + 4 + 1
        .Top = mdiRect.Bottom - dropZone \ Screen.TwipsPerPixelX - 4
        .Right = rightDock.Right - dropZone \ Screen.TwipsPerPixelX
        .Bottom = mdiRect.Bottom - 4
    End With
    
    'See if the top/left corner of this form is in Picture1's screen rectangle
    'As we have set RepositionForm to false, we are responsible for positioning the form
    If (Not bDocked) Then
        If PtInRect(leftDock, ptx, pty) Then
            Dock_Left
        ElseIf PtInRect(rightDock, ptx, pty) Then
            Dock_Right
        ElseIf PtInRect(topDock, ptx, pty) Then
            Dock_Top
        ElseIf PtInRect(botDock, ptx, pty) Then
            Dock_Bottom
        End If
        
    Else
        Select Case (frmMDI.Picture1.Align)
            Case 3
                If (PtInRect(picRect, ptx, pty)) Then
                    Dock_Left
                Else
                
                    If PtInRect(rightDock, ptx, pty) Then
                        Dock_Right
                    ElseIf PtInRect(topDock, ptx, pty) Then
                        Dock_Top
                    ElseIf PtInRect(botDock, ptx, pty) Then
                        Dock_Bottom
                    Else
                        UnDock
                    End If
                    
                End If ' (PtInRect(picRect, ptx, pty))
        
            Case 4
                If (PtInRect(picRect, ptx, pty)) Then
                    Dock_Right
                Else
                                
                    If PtInRect(leftDock, ptx, pty) Then
                        Dock_Left
                    ElseIf PtInRect(topDock, ptx, pty) Then
                        Dock_Top
                    ElseIf PtInRect(botDock, ptx, pty) Then
                        Dock_Bottom
                    Else
                        UnDock
                    End If
                
                End If ' (PtInRect(picRect, ptx, pty))
                
            Case 1
                If (PtInRect(picRect, ptx, pty)) Then
                    Dock_Top
                Else
                                
                    If PtInRect(leftDock, ptx, pty) Then
                        Dock_Left
                    ElseIf PtInRect(rightDock, ptx, pty) Then
                        Dock_Right
                    ElseIf PtInRect(botDock, ptx, pty) Then
                        Dock_Bottom
                    Else
                        UnDock
                    End If
                
                End If ' (PtInRect(picRect, ptx, pty))
                
            Case 2
                If (PtInRect(picRect, ptx, pty)) Then
                    Dock_Bottom
                Else
                    
                    If PtInRect(leftDock, ptx, pty) Then
                        Dock_Left
                    ElseIf PtInRect(rightDock, ptx, pty) Then
                        Dock_Right
                    ElseIf PtInRect(topDock, ptx, pty) Then
                        Dock_Top
                    Else
                        UnDock
                    End If
                    
                End If ' (PtInRect(picRect, ptx, pty))
                
        End Select ' Case (frmMDI.Picture1.Align)

    End If ' (Not bDocked
    ' Reset the moving flag and store the form dimensions
    bMoving = False
    StoreFormDimensions

End Sub
'*
Private Sub StoreFormDimensions()

   'Store the height/width values
    If Not bMoving Then
        If bDocked Then
            lDockedWidth = Me.Width
            lDockedHeight = Me.Height
        Else
            lFloatingLeft = Me.Left
            lFloatingTop = Me.Top
            lFloatingWidth = Me.Width
            lFloatingHeight = Me.Height
        End If
    End If
    
End Sub
'*
Private Sub DrawDragRectangle(ByVal x As Long, ByVal Y As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal lWidth As Long)

    'Draw a rectangle using the Win32 API

    Dim hDC As Long
    Dim hPen As Long
    hPen = CreatePen(PS_SOLID, lWidth, &HE0E0E0)
    hDC = GetDC(0)
    Call SelectObject(hDC, hPen)
    Call SetROP2(hDC, R2_NOTXORPEN)
    Call Rectangle(hDC, x, Y, X1, Y1)
    Call SelectObject(hDC, GetStockObject(BLACK_PEN))
    Call DeleteObject(hPen)
    Call SelectObject(hDC, hPen)
    Call ReleaseDC(0, hDC)
    
End Sub
'*
Private Sub DrawRect(rct As Rect)

    With rct

        'Draw a rectangle using the Win32 API
    
        Dim hDC As Long
        Dim hPen As Long
        hPen = CreatePen(PS_SOLID, 3, &HE0E0E0)
        hDC = GetDC(0)
        Call SelectObject(hDC, hPen)
        Call SetROP2(hDC, R2_NOTXORPEN)
        Call Rectangle(hDC, .Left, .Top, .Right, .Bottom)
        Call SelectObject(hDC, GetStockObject(BLACK_PEN))
        Call DeleteObject(hPen)
        Call SelectObject(hDC, hPen)
        Call ReleaseDC(0, hDC)
    
    End With
    
End Sub
'*
Private Sub Form_Moved(ptx As Long, pty As Long)

    Dim formRect As Rect
    Dim mdiRect As Rect
    Dim picRect As Rect
    Dim leftDock As Rect
    Dim rightDock As Rect
    Dim topDock As Rect
    Dim botDock As Rect
    
    'Set the moving flag so we dont store the wrong dimensions
    bMoving = True
    
    'If over Picture1 on frmMDI which we are using as a Dock, change the width to that of
    'Picture1, else change it to the 'floating width and height
    'For the purposes of this example, we only dock if the top left corner
    'of this form is within the area bounded by Picture1
    
    ' Get the screen based coordinates of our MDI Form
    GetWindowRect frmMDI.hWnd, mdiRect
    
    ' Get the screen based coordinates of our PictureBox
    GetWindowRect frmMDI.Picture1.hWnd, picRect
    
    ' Set up the drop zone regions. These will be used for
    ' check to see if the form is to be docked or not.
    
    With leftDock
        .Left = mdiRect.Left + 4
        .Top = mdiRect.Top + TitleBarHeight '(mdiRect.Bottom - mdiRect.Top - frmMDI.ScaleHeight \ Screen.TwipsPerPixelY) - 8
        .Right = .Left + dropZone \ Screen.TwipsPerPixelX
        .Bottom = .Top + frmMDI.ScaleHeight \ Screen.TwipsPerPixelY + 4
    End With
    
    With rightDock
        .Left = mdiRect.Right - dropZone \ Screen.TwipsPerPixelX - 4
        .Top = mdiRect.Top + TitleBarHeight '(mdiRect.Bottom - mdiRect.Top - frmMDI.ScaleHeight \ Screen.TwipsPerPixelY) - 8
        .Right = mdiRect.Right - 4
        .Bottom = .Top + frmMDI.ScaleHeight \ Screen.TwipsPerPixelY + 4
    End With
    
    With topDock
        .Left = mdiRect.Left + dropZone \ Screen.TwipsPerPixelX + 4 + 1
        .Top = mdiRect.Top + TitleBarHeight '(mdiRect.Bottom - mdiRect.Top - frmMDI.ScaleHeight \ Screen.TwipsPerPixelY) - 8
        .Right = rightDock.Right - dropZone \ Screen.TwipsPerPixelX
        .Bottom = .Top + dropZone \ Screen.TwipsPerPixelX
    End With
    
    With botDock
        .Left = mdiRect.Left + dropZone \ Screen.TwipsPerPixelX + 4 + 1
        .Top = mdiRect.Bottom - dropZone \ Screen.TwipsPerPixelX - 4
        .Right = rightDock.Right - dropZone \ Screen.TwipsPerPixelX
        .Bottom = mdiRect.Bottom - 4
    End With

    'DrawRect leftDock
    'DrawRect rightDock
    'DrawRect topDock
    'DrawRect botDock
    
    'Debug.Print "a) "; mdiRect.Top; " <--> "; mdiRect.Bottom
    'Debug.Print "b) "; topDock.Top; " <-->"; topDock.Bottom
    
    'See if the top/left corner of this form is in Picture1's screen rectangle
    
    If (Not bDocked) Then
        If PtInRect(leftDock, ptx, pty) Then
            Calc_Left ptx
        ElseIf PtInRect(rightDock, ptx, pty) Then
            Calc_Right ptx
        ElseIf PtInRect(topDock, ptx, pty) Then
            Calc_Top ptx
        ElseIf PtInRect(botDock, ptx, pty) Then
            Calc_Bottom ptx
        Else
            Calc_Default
        End If
        
    Else
        Select Case (frmMDI.Picture1.Align)
            Case 3
                If (PtInRect(picRect, ptx, pty)) Then
                    Calc_Left ptx
                Else
                
                    If PtInRect(rightDock, ptx, pty) Then
                        Calc_Right ptx
                    ElseIf PtInRect(topDock, ptx, pty) Then
                        Calc_Top ptx
                    ElseIf PtInRect(botDock, ptx, pty) Then
                        Calc_Bottom ptx
                    Else
                        Calc_Default
                    End If
                    
                End If ' (PtInRect(picRect, ptx, pty))
        
            Case 4
                If (PtInRect(picRect, ptx, pty)) Then
                    Calc_Right ptx
                Else
                                
                    If PtInRect(leftDock, ptx, pty) Then
                        Calc_Left ptx
                    ElseIf PtInRect(topDock, ptx, pty) Then
                        Calc_Top ptx
                    ElseIf PtInRect(botDock, ptx, pty) Then
                        Calc_Bottom ptx
                    Else
                        Calc_Default
                    End If
                
                End If ' (PtInRect(picRect, ptx, pty))
                
            Case 1
                If (PtInRect(picRect, ptx, pty)) Then
                    Calc_Top ptx
                Else
                                
                    If PtInRect(leftDock, ptx, pty) Then
                        Calc_Left ptx
                    ElseIf PtInRect(rightDock, ptx, pty) Then
                        Calc_Right ptx
                    ElseIf PtInRect(botDock, ptx, pty) Then
                        Calc_Bottom ptx
                    Else
                        Calc_Default
                    End If
                
                End If ' (PtInRect(picRect, ptx, pty))
                
            Case 2
                If (PtInRect(picRect, ptx, pty)) Then
                    Calc_Bottom ptx
                Else
                    
                    If PtInRect(leftDock, ptx, pty) Then
                        Calc_Left ptx
                    ElseIf PtInRect(rightDock, ptx, pty) Then
                        Calc_Right ptx
                    ElseIf PtInRect(topDock, ptx, pty) Then
                        Calc_Top ptx
                    Else
                        Calc_Default
                    End If
                    
                End If ' (PtInRect(picRect, ptx, pty))
                
        End Select ' Case (frmMDI.Picture1.Align)

    End If ' (Not bDocked)


    
End Sub
'*
Sub UnDock()

    ' If it was docked before, undock it
    Call SetParent(Me.hWnd, dockParent)
    Me.Visible = False
    bDocked = False
    frmMDI!Picture1.Visible = False
    Me.Visible = True
    Call SendMessage(Me.hWnd, WM_NCACTIVATE, 1, 0)
        
End Sub

