VERSION 5.00
Begin VB.Form frmMultiClip 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3180
   ClientLeft      =   9105
   ClientTop       =   1740
   ClientWidth     =   2745
   Icon            =   "frmMultiClip.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMultiClip.frx":0E42
   ScaleHeight     =   3180
   ScaleWidth      =   2745
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
      Height          =   3180
      Left            =   0
      ScaleHeight     =   3180
      ScaleWidth      =   330
      TabIndex        =   1
      Top             =   0
      Width           =   330
   End
   Begin VB.ListBox lstClip 
      Height          =   3180
      Left            =   330
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   2415
   End
   Begin VB.Menu mnuMultiClip 
      Caption         =   "Edit"
      Visible         =   0   'False
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnuSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "frmMultiClip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim xpos As Long
Dim ypos As Long


Private Sub Form_Load()
    Call rotateText("E-Type MultiClip", picBorder)
End Sub

Private Sub lstClip_DblClick()

    ' Loop through the items in the ListBox control.
    For ClipItem = 0 To lstClip.ListCount - 1
    
        ' If the item is selected...
        If lstClip.Selected(ClipItem) = True Then
        
            '...paste the Selected item.
            frmMDI.ActiveForm.ActiveControl.SelText = ""    'This step is crucial!!! for undoing actions
            ' Place the text from the Clipboard into the active control.
            frmMDI.ActiveForm.ActiveControl.SelText = lstClip.List(ClipItem)
            ' Set focus back to the active window
            frmMDI.ActiveForm.ActiveControl.SetFocus
          
      End If
    
    Next ClipItem

End Sub

Private Sub lstClip_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    ' Show popupmenu
    If Button = vbRightButton Then 'do the popup menu
        PopupMenu mnuMultiClip
    End If

End Sub

Private Sub lstClip_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    ' testing wich line the mousepionter is over and
    ' displaying the text in the line as tooltip
    
    Dim lXPoint As Long
    Dim lYPoint As Long
    Dim lIndex As Long
    Dim tooltiptxt As String
    
    If Button = 0 Then ' if no button was pressed
        lXPoint = CLng(x / Screen.TwipsPerPixelX)
        lYPoint = CLng(y / Screen.TwipsPerPixelY)
        With lstClip
            ' get selected item from list
            lIndex = SendMessage(.hwnd, LB_ITEMFROMPOINT, 0, ByVal ((lYPoint * 65536) + lXPoint))
            
            ' show tip or clear last one
            If (lIndex >= 0) And (lIndex <= .ListCount) Then
                tooltiptxt = .List(lIndex)
                tooltiptxt = Left(tooltiptxt, 80) 'max 40 caracters long
                .ToolTipText = tooltiptxt & "...."
            Else
                .ToolTipText = ""
            End If
        End With
    End If

End Sub

Private Sub mnuClose_Click()
    Unload frmMultiClip
End Sub

Private Sub mnuPaste_Click()
    lstClip_DblClick
End Sub

Public Sub KillDupes()
    ' this sub kills all duplicate lines in lstClip

    Dim x As String
    Dim xa As String
    Dim xaa As String
    Dim xx As String
    Dim i

    On Error GoTo lol
    For i = 0 To lstClip.ListCount - 1
        DoEvents
        lstClip.ListIndex = i
        x = lstClip.List(i)
        xx = lstClip.List(i + 1)
        'trims all spaces in a the current item
        xa = trimtext(x)
        xaa = trimtext(xx)
        'if dupe is found removes it
        If LCase(xa) = LCase(xaa) Then
            DoEvents
            lstClip.RemoveItem i
            i = i - 1
        End If
    Next i
    Exit Sub
lol:
    Exit Sub

End Sub

Function trimtext(txt As String) As String
    'This funktion is used by KillDupes on frmMultiKlipp
    Dim i
    Dim xx
    Dim x
    'starts from the beginging to the end of the text
    For i = 1 To Len(txt)
        DoEvents
        'checks the letters one by one for a space
        x = Mid(txt, i)
        x = Left(x, 1)
        If x = " " Then
            Else
            xx = xx + x
        End If
    Next i
    trimtext = xx
End Function

