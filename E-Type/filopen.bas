Attribute VB_Name = "modFileIO"
'*** Standard module with procedures for working with   ***
'*** files. Part of the MDI Notepad sample application. ***
'**********************************************************
Option Explicit

Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

' used by BrowseForFolder start
Public Type BrowseInfo
     hwndOwner As Long
     pIDLRoot As Long
     pszDisplayName As Long
     lpszTitle As Long
     ulFlags As Long
     lpfnCallback As Long
     lParam As Long
     iImage As Long
End Type
Public Const BIF_RETURNONLYFSDIRS = 1
Public Const MAX_PATH = 260
Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
' used by BrowseForFolder end

Sub FileOpenProc()
    
    Dim intRetVal
    On Error Resume Next
    Dim strOpenFileName As String
    frmMDI.CMDialog1.DialogTitle = "Open File"
    frmMDI.CMDialog1.FileName = ""
    frmMDI.CMDialog1.Filter = "Files (*.txt)|*.txt|All Files (*.*)|*.*"
    frmMDI.CMDialog1.FilterIndex = 1
    frmMDI.CMDialog1.DefaultExt = "txt"
    frmMDI.CMDialog1.ShowOpen
    If Err <> 32755 Then    ' User chose Cancel.
        strOpenFileName = frmMDI.CMDialog1.FileName
        
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

    End If
End Sub

Function GetFileName(FileName As Variant)
    ' Display a Save As dialog box and return a filename.
    ' If the user chooses Cancel, return an empty string.
    On Error Resume Next
    frmMDI.CMDialog1.DialogTitle = "Save File"
    frmMDI.CMDialog1.FileName = FileName
    frmMDI.CMDialog1.Filter = "Files (*.txt)|*.txt|All Files (*.*)|*.*"
    frmMDI.CMDialog1.FilterIndex = 1
    frmMDI.CMDialog1.ShowSave
    If Err <> 32755 Then    ' User chose Cancel.
        GetFileName = frmMDI.CMDialog1.FileName
    Else
        GetFileName = ""
    End If
End Function

Function OnRecentFilesList(FileName) As Integer
  Dim i         ' Counter variable.

  For i = 1 To 4
    If frmMDI.mnuRecentFile(i).Caption = FileName Then
      OnRecentFilesList = True
      Exit Function
    End If
  Next i
    OnRecentFilesList = False
End Function

Sub OpenFile(FileName)
'    Dim fIndex As Integer
    
    On Error Resume Next
    ' Open the selected file.
    Open FileName For Input As #1
    If Err Then
        MsgBox "Can't open file: " + FileName
        Exit Sub
    End If
    ' Change the mouse pointer to an hourglass.
    Screen.MousePointer = 11
    
    ' Change the form's caption and display the new text.
    fIndex = FindFreeIndex()
    Document(fIndex).Tag = fIndex
    Document(fIndex).Caption = FileName    'UCase(FileName)
    Call frmMDI.ActiveForm.SetFonts
    Document(fIndex).Text1.LoadFile (FileName)
    Document(fIndex).Text1.Tag = "docUnChanged"
    Document(fIndex).Show
    Close #1
    ' Reset the mouse pointer.
    Screen.MousePointer = 0
    
End Sub

Sub SaveFileAs(FileName)
    On Error Resume Next
    Dim strContents As String

    ' Open the file.
    Open FileName For Output As #1
    ' Place the contents of the notepad into a variable.
    strContents = frmMDI.ActiveForm.Text1.text
    ' Display the hourglass mouse pointer.
    Screen.MousePointer = 11
    ' Write the variable contents to a saved file.
    Print #1, strContents
    Close #1
    ' Reset the mouse pointer.
    Screen.MousePointer = 0
    ' Set the form's caption.
    If Err Then
        MsgBox Error, 48, App.Title
    Else
        frmMDI.ActiveForm.Caption = FileName    'UCase(FileName)
        ' updating the doctab text and tooltip
        frmMDI.tabDocuments.Tabs("key" & frmMDI.ActiveForm.Tag).ToolTipText = frmMDI.ActiveForm.Caption
        frmMDI.tabDocuments.Tabs("key" & frmMDI.ActiveForm.Tag).Caption = StripPath(frmMDI.ActiveForm.Caption)
        ' Reset the dirty flag.
        frmMDI.ActiveForm.Text1.Tag = "docUnchanged"
    End If
End Sub

Sub UpdateFileMenu(FileName)
        Dim intRetVal As Integer
        ' Check if the open filename is already in the File menu control array.
        intRetVal = OnRecentFilesList(FileName)
        If Not intRetVal Then
            ' Write open filename to the registry.
            WriteRecentFiles (FileName)
        End If
        ' Update the list of the most recently opened files in the File menu control array.
        GetRecentFiles
End Sub

 Sub InsertFile()
    ' This procedure opens the filedialog and inserts the choosen file
    ' at the cursor possition
   
    Dim choice As Integer
    Dim filenum As Integer
    On Error GoTo inserterrortrap
    frmMDI.CMDialog1.DialogTitle = "Insert File"
    frmMDI.CMDialog1.FileName = ""
    frmMDI.CMDialog1.ShowOpen
    If frmMDI.CMDialog1.FileName <> "" Then
        Screen.MousePointer = 11
        filenum = FreeFile
        Open frmMDI.CMDialog1.FileName For Input As filenum
             frmMDI.ActiveForm.ActiveControl.SelText = Input(LOF(filenum), filenum)
        Close (filenum)
        Screen.MousePointer = 0
    End If
    Exit Sub
inserterrortrap:
    Screen.MousePointer = 0
    MsgBox "Error opening " & frmMDI.CMDialog1.FileName, vbExclamation, "Error!"
    Exit Sub
    Resume Next

End Sub

Public Function BrowseForFolder(hwndOwner As Long, sPrompt As String) As String
     
    'declare variables to be used
     Dim iNull As Integer
     Dim lpIDList As Long
     Dim lResult As Long
     Dim spath As String
     Dim udtBI As BrowseInfo

    'initialise variables
     With udtBI
        .hwndOwner = hwndOwner
        .lpszTitle = lstrcat(sPrompt, "")
        .ulFlags = BIF_RETURNONLYFSDIRS
     End With

    'Call the browse for folder API

     lpIDList = SHBrowseForFolder(udtBI)
     
    'get the resulting string path
     If lpIDList Then
        spath = String$(MAX_PATH, 0)
        lResult = SHGetPathFromIDList(lpIDList, spath)
        Call CoTaskMemFree(lpIDList)
        iNull = InStr(spath, vbNullChar)
        If iNull Then spath = Left$(spath, iNull - 1)
     End If

    'If cancel was pressed, sPath = ""
     BrowseForFolder = spath

End Function

Public Function CompareFiles(ByVal MasterFile As String, ByVal SourceFile As String, ByVal NewFile As String) As String
     '
     ' Purpose: Allows a efficent means under VB6 to manage
     ' duplicate records in text files. Logic: Anything found in a
     ' SourceFile NOT found in MasterFile is written to NewFile.
     ' NewFile will contain only new lines that do not exist in
     ' MasterFile. NewFile is Appended, so should be deleted after
     ' a Compare operation. Also, SourceFile should be Appended so
     ' MasterFile to aid in tracking duplicate data if desired
     ' over time.
     '
     ' Requires: Microsoft Scripting runtime (scrrun.dll).
     ' Add using Project Reference.
     ' Usage Example:
     ' CompareFiles "c:\master.txt", "c:\source.txt", "c:\new.txt"

     Dim oFileSystem As FileSystemObject
     Dim oNewFile As TextStream
     Dim oMasterFile As TextStream
     Dim oSourceFile As TextStream
     Dim sCurrent
     Set oFileSystem = New FileSystemObject
     Set oMasterFile = oFileSystem.OpenTextFile(MasterFile, ForReading)
     Set oSourceFile = oFileSystem.OpenTextFile(SourceFile, ForReading)
     Set oNewFile = oFileSystem.OpenTextFile(NewFile, ForAppending, True)
     MasterFile = oMasterFile.ReadAll

     Do Until oSourceFile.AtEndOfStream
         sCurrent = oSourceFile.ReadLine
         If InStr(1, MasterFile, sCurrent) Then
             ' Line currently already exists in MasterFile (ignore it)
         Else
             oNewFile.WriteLine sCurrent ' Write what does not exists in MasterFile
         End If
     Loop
     oNewFile.Close
     oMasterFile.Close
     oSourceFile.Close
     Set oSourceFile = Nothing
     Set oMasterFile = Nothing
     Set oNewFile = Nothing
     Set oFileSystem = Nothing
 End Function


Public Function GetTmpPath()
    ' If you have a project that produces temporary files during runtime, it is
    ' good practice to make these files in the official Windows temporary
    ' directory.  This directory is usually "c:\windows\temp", although it can vary
    ' with different flavours of windows.  You can retrieve the path to the
    ' temporary directory by using the following code.

    ' An example of how to use the function:
    '  Call MsgBox("The temp path is " & GetTmpPath, vbInformation)
    
    Dim strFolder As String
    Dim lngResult As Long
    strFolder = String(MAX_PATH, 0)
    lngResult = GetTempPath(MAX_PATH, strFolder)
    If lngResult <> 0 Then
        GetTmpPath = Left(strFolder, InStr(strFolder, _
        Chr(0)) - 1)
    Else
        GetTmpPath = ""
    End If
End Function

Public Function FileExists(ByVal PathName As String) As Boolean
'    Add two textboxes to form (Text1 and Text2)
'    Add a Commandbutton to form.
'
'    Paste following code:
'
'    Private Sub Command1_Click()
'     Text2 = FileExists(Text1)
'    End Sub
'
'    In Text1 type a filename (full path) You know exits.
'    I.e. C:\config.sys, press Command1
'    Text2 should now show 'true'.
    
    FileExists = IIf(Dir$(PathName) = "", False, True)

End Function


