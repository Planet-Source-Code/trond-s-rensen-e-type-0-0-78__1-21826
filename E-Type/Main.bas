Attribute VB_Name = "modMain"
'*** Global module for E-Type.  ***
'**********************************
Option Explicit

' User-defined type to store information about child forms
Type FormState
    Deleted As Boolean
    dirty As Boolean
    Color As Long
End Type

Public FState() As FormState            ' Array of user-defined types
Public Document() As New frmNotePad     ' Array of child form objects
Public gFindString As String            ' Holds the search text.
Public gFindCase As Integer             ' Key for case sensitive search
Public gFindDirection As Integer        ' Key for search direction.
Public gCurPos As Integer               ' Holds the cursor location.
Public gFirstTime As Integer            ' Key for start position.
Public gLeftMargin As Integer           ' Print Preview
Public gRightMargin As Integer          ' Print Preview
Public gTopMargin As Integer            ' Print Preview
Public gBottomMargin As Integer         ' Print Preview
Public gprint As Boolean                ' Print Preview
Public iLine As Double                  ' linenumbers - total lines
Public cLine As Double                  ' linenumbers - current line
Public vLine As Double                  ' linenumbers - ? lines
Public fIndex As Integer                ' file index
Global ClipItem  As Integer             ' The selected item in cliplist on frmMultiClip
Global ftime As SYSTEMTIME              ' file info for the current open file
Global tfilename As String              ' file info for the current open file
Global filedata As WIN32_FIND_DATA      ' file info for the current open file
Global CancelDirSearch As Boolean       ' Used to stop filesearch

Public Const ThisApp = "E-Type"         ' Registry App constant.
Public Const ThisKey = "Recent Files"   ' Registry Key constant.
Global Const LB_ITEMFROMPOINT = &H1A9   ' Multiclip lineinfo from mousepointer

Public Const EM_GETSEL As Double = &HB0     ' Linenumbers
Public Const EM_SETSEL As Long = &HB1       ' Linenumbers
Public Const EM_GETLINECOUNT As Long = &HBA ' Linenumbers
Public Const EM_LINEINDEX As Long = &HBB    ' Linenumbers
Public Const EM_LINELENGTH As Long = &HC1   ' Linenumbers
Public Const EM_LINEFROMCHAR As Long = &HC9 ' Linenumbers
Public Const EM_SCROLLCARET As Long = &HB7  ' Linenumbers
Public Const WM_SETREDRAW As Long = &HB     ' Linenumbers
Public Const EM_GETLINE = &HC4              ' Linenumbers
Public Const WM_USER = &H400
Public Const EM_HIDESELECTION = WM_USER + 63
Public Const CB_SHOWDROPDOWN = WM_USER + 15
Public Const MAX_PATH = 260
Public Const INVALID_HANDLE_VALUE = -1

Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOP = 0
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Const FILE_ATTRIBUTE_READONLY = &H1
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_TEMPORARY = &H100
Public Const FILE_ATTRIBUTE_COMPRESSED = &H800
Public Const EM_GETFIRSTVISIBLELINE = &HCE 'picLines lines

Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal Cx As Long, ByVal Cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetFocus Lib "user32" () As Long
Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
'Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
'Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Type FILETIME
   LowDateTime          As Long
   HighDateTime         As Long
End Type

' denne fulgte med fileinfo
Type WIN32_FIND_DATA
   dwFileAttributes     As Long
   ftCreationTime       As FILETIME
   ftLastAccessTime     As FILETIME
   ftLastWriteTime      As FILETIME
   nFileSizeHigh        As Long
   nFileSizeLow         As Long
   dwReserved0          As Long
   dwReserved1          As Long
   cFileName            As String * 260  'MUST be set to 260
   cAlternate           As String * 14
End Type

Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Public Enum LineInfo
    [Line count] = 0
    [Cursor Position] = 1
    [Current Line Number] = 2
    [Current Line Start] = 3
    [Current Line End] = 4
    [Current Line Length] = 5
    [Current Line Cursor Position] = 6
    [Line Start] = 7
    [Line End] = 8
    [Line Length] = 9
End Enum

Function AnyPadsLeft() As Integer
    Dim i As Integer        ' Counter variable

    ' Cycle through the document array.
    ' Return true if there is at least one open document.
    For i = 1 To UBound(Document)
        If Not FState(i).Deleted Then
            AnyPadsLeft = True
            Exit Function
        End If
    Next
End Function

Sub EditCopyProc()
    '********** ikke i bruk ****************
    ' Copy the selected text onto the Clipboard.
    Clipboard.SetText frmMDI.ActiveForm.ActiveControl.SelText
    ' Copy the selected text to the list on frmMultiClip
    frmMultiClip.lstClip.AddItem (Clipboard.GetText)
    Call frmMultiClip.KillDupes

End Sub

Sub EditCutProc()
    '********** ikke i bruk ****************
    ' Copy the selected text onto the Clipboard.
    Clipboard.SetText frmMDI.ActiveForm.ActiveControl.SelText
    ' Copy the selected text to the list on frmMultiClip
    frmMultiClip.lstClip.AddItem (Clipboard.GetText)
    ' Delete the selected text.
    frmMDI.ActiveForm.ActiveControl.SelText = ""
End Sub

Sub EditPasteProc()
    '********** ikke i bruk ****************
    ' Place the text from the Clipboard into the active control.
    frmMDI.ActiveForm.ActiveControl.SelText = Clipboard.GetText()
End Sub

Sub FileNew()

    ' Find the next available index and show the child form.
    fIndex = FindFreeIndex()
    Document(fIndex).Tag = fIndex
    Document(fIndex).Caption = "Untitled: " & Document(fIndex).Tag
    Document(fIndex).Show

End Sub

Function FindFreeIndex() As Integer
    Dim i As Integer
    Dim ArrayCount As Integer

    ArrayCount = UBound(Document)

    ' Cycle through the document array. If one of the
    ' documents has been deleted, then return that index.
    For i = 1 To ArrayCount
        If FState(i).Deleted Then
            FindFreeIndex = i
            FState(i).Deleted = False
            Exit Function
        End If
    Next

    ' If none of the elements in the document array have
    ' been deleted, then increment the document and the
    ' state arrays by one and return the index to the
    ' new element.
    ReDim Preserve Document(1 To ArrayCount + 1)
    ReDim Preserve FState(1 To ArrayCount + 1)
    FindFreeIndex = UBound(Document)
End Function

Sub FindIt()
    Dim intStart As Integer
    Dim intPos As Integer
    Dim strFindString As String
    Dim strSourceString As String
    Dim strMsg As String
    Dim intResponse As Integer
    Dim intOffset As Integer
    
    ' Set offset variable based on cursor position.
    If (gCurPos = frmMDI.ActiveForm.ActiveControl.SelStart) Then
        intOffset = 1
    Else
        intOffset = 0
    End If

    ' Read the public variable for start position.
    If gFirstTime Then intOffset = 0
    ' Assign a value to the start value.
    intStart = frmMDI.ActiveForm.ActiveControl.SelStart + intOffset
        
    ' If not case sensitive, convert the string to upper case
    If gFindCase Then
        strFindString = gFindString
        strSourceString = frmMDI.ActiveForm.ActiveControl.text
    Else
        strFindString = UCase(gFindString)
        strSourceString = UCase(frmMDI.ActiveForm.ActiveControl.text)
    End If
            
    ' Search for the string.
    If gFindDirection = 1 Then
        intPos = InStr(intStart + 1, strSourceString, strFindString)
    Else
        For intPos = intStart - 1 To 0 Step -1
            If intPos = 0 Then Exit For
            If Mid(strSourceString, intPos, Len(strFindString)) = strFindString Then Exit For
        Next
    End If

    ' If the string is found...
    If intPos Then
        frmMDI.ActiveForm.ActiveControl.SelStart = intPos - 1
        frmMDI.ActiveForm.ActiveControl.SelLength = Len(strFindString)
    Else
        strMsg = "E-Type has finiched searching the document. No more occurrences of " & Chr(34) & gFindString & Chr(34) & " found."
        intResponse = MsgBox(strMsg, vbExclamation, App.Title)
    End If
    
    ' Reset the public variables
    gCurPos = frmMDI.ActiveForm.ActiveControl.SelStart
    gFirstTime = False
End Sub

Sub GetRecentFiles()
    ' This procedure demonstrates the use of the GetAllSettings function,
    ' which returns an array of values from the Windows registry. In this
    ' case, the registry contains the files most recently opened.  Use the
    ' SaveSetting statement to write the names of the most recent files.
    ' That statement is used in the WriteRecentFiles procedure.
    Dim i, j As Integer
    Dim varFiles As Variant ' Varible to store the returned array.
    
    ' Get recent files from the registry using the GetAllSettings statement.
    ' ThisApp and ThisKey are constants defined in this module.
    If GetSetting(ThisApp, ThisKey, "RecentFile1") = Empty Then Exit Sub
    
    varFiles = GetAllSettings(ThisApp, ThisKey)
    
    For i = 0 To UBound(varFiles, 1)
        frmMDI.mnuRecentFile(0).Visible = True
        frmMDI.mnuRecentFile(i).Caption = varFiles(i, 1)
        frmMDI.mnuRecentFile(i).Visible = True
            ' Iterate through all the documents and update each menu.
            For j = 1 To UBound(Document)
                If Not FState(j).Deleted Then
                    Document(j).mnuRecentFile(0).Visible = True
                    Document(j).mnuRecentFile(i + 1).Caption = varFiles(i, 1)
                    Document(j).mnuRecentFile(i + 1).Visible = True
                End If
            Next j
    Next i
End Sub

Sub WriteRecentFiles(OpenFileName)
    ' This procedure uses the SaveSettings statement to write the names of
    ' recently opened files to the System registry. The SaveSetting
    ' statement requires three parameters. Two of the parameters are
    ' stored as constants and are defined in this module.  The GetAllSettings
    ' function is used in the GetRecentFiles procedure to retrieve the
    ' file names stored in this procedure.
    
    Dim i, j As Integer
    Dim strFile, key As String

    ' Copy RecentFile1 to RecentFile2, and so on.
    For i = 3 To 1 Step -1
        key = "RecentFile" & i
        strFile = GetSetting(ThisApp, ThisKey, key)
        If strFile <> "" Then
            key = "RecentFile" & (i + 1)
            SaveSetting ThisApp, ThisKey, key, strFile
        End If
    Next i
  
    ' Write the open file to first recent file.
    SaveSetting ThisApp, ThisKey, "RecentFile1", OpenFileName
End Sub

Public Sub GetFileStats()
    
    ' This procedure displays fileinfo on the status bar
    ' It calls the FindFile procedure to get the data
    
    tfilename = frmMDI.ActiveForm.Caption

    filedata = Findfile(tfilename)        ' Get information
    
    If filedata.nFileSizeHigh = 0 Then    ' Put size into text box
      frmMDI.SBarMain.Panels(5).text = "File Size: " & filedata.nFileSizeLow & " Bytes"
    Else
      frmMDI.SBarMain.Panels(5).text = "File Size: " & filedata.nFileSizeHigh & "Bytes"
    End If
    
    ' Do not change the order on the next lines!!
    Call FileTimeToSystemTime(filedata.ftLastWriteTime, ftime)  ' Determine Last Modified date and time
    If ftime.wDay = "1" And ftime.wMonth = "1" And ftime.wYear = "1601" Then
        frmMDI.SBarMain.Panels(4).text = "Mod: " & Format(Now, "d/m/yyyy h:mm:ss ")
    Else
        frmMDI.SBarMain.Panels(4).text = "Mod: " & ftime.wDay & "." & ftime.wMonth & "." & ftime.wYear & " " & ftime.wHour & ":" & ftime.wMinute & ":" & ftime.wSecond
    End If
End Sub

Public Sub SetStatusBar(strStatusText As String, Optional strPanel As String = "Status")
    
    ' This sub puts the data from strStatusText in to the statusbar
    ' strStatusText contains data about the cursor possition,
    ' lines in doc and corrent line
    
    frmMDI.SBarMain.Panels(2).text = strStatusText

End Sub

'************************************************************
'Builds the button menu for the window toolbar button
'************************************************************
Public Sub RebuildWinList()
    
'    Dim i As Long
'
'    frmMDI.tbToolBar.Buttons("SelWindow").ButtonMenus.Clear
'    For i = 1 To Forms.Count - 1
'        frmMDI.tbToolBar.Buttons("SelWindow").ButtonMenus.Add i, , Forms(i).Caption
'    Next
    Dim i As Long
    frmMDI.tbToolBar.Buttons("SelWindow").ButtonMenus.Clear
   ' frmMDI.tbToolBar.Buttons("selWindow").ButtonMenus.Clear
    
    For i = 1 To Forms.Count - 1
        frmMDI.tbToolBar.Buttons("SelWindow").ButtonMenus.Add i, , Forms(i).Caption
    Next

End Sub

Function StripPath(T) As String
' This function strips the filename out of a full path

Dim x As Integer
Dim ct As Integer
    StripPath = T
    x = InStr(T, "\")
    Do While x
        ct = x
        x = InStr(ct + 1, T, "\")
    Loop
    If ct > 0 Then StripPath = Mid(T, ct + 1)
    
End Function

Public Function getLineInfo(txtObj As Object, info As LineInfo, Optional lineNumber As Long) As Double 'Long

   '***************************************************************
   ' Name: text object line info
   ' Description:The first function returns usefull information
   ' about text box objects.
   ' These include:
   
   ' [Line count] = 0
   ' [Cursor Position] = 1
   ' [Current Line Number] = 2
   ' [Current Line Start] = 3
   ' [Current Line End] = 4
   ' [Current Line Length] = 5
   ' [Current Line Cursor Position] = 6
   ' [Line Start] = 7
   ' [Line End] = 8
   ' [Line Length] = 9
    
    
    Dim cursorPoint As Long
    'Record where the cursor is
    
    cursorPoint = txtObj.SelStart


    Select Case info
        Case Is = 0 ' = "lineCount"
        getLineInfo = SendMessageLong(txtObj.hwnd, EM_GETLINECOUNT, 0, 0&)
        Case Is = 1 ' = "cursorPosition"
        getLineInfo = (SendMessageLong(txtObj.hwnd, EM_GETSEL, 0, 0&) \ &H10000) + 1
        Case Is = 2 ' = "currentLineNumber"
        getLineInfo = (SendMessageLong(txtObj.hwnd, EM_LINEFROMCHAR, -1, 0&)) + 1
        Case Is = 3 ' = "currentLineStart"
        getLineInfo = SendMessageLong(txtObj.hwnd, EM_LINEINDEX, -1, 0&) + 1
        Case Is = 4 ' = "currentLineEnd"
        getLineInfo = SendMessageLong(txtObj.hwnd, EM_LINEINDEX, -1, 0&) + 1 + SendMessageLong(txtObj.hwnd, EM_LINELENGTH, -1, 0&)
        Case Is = 5 ' = "currentLineLength"
        getLineInfo = SendMessageLong(txtObj.hwnd, EM_LINELENGTH, -1, 0&)
        Case Is = 6 ' = "currentLineCursorPosition"
        getLineInfo = (SendMessageLong(txtObj.hwnd, EM_GETSEL, 0, 0&) \ &H10000) + 1 - SendMessageLong(txtObj.hwnd, EM_LINEINDEX, getLineInfo(txtObj, [Current Line Number]) - 1, 0&)
        Case Is = 7 ' = "lineStart"
        getLineInfo = (SendMessageLong(txtObj.hwnd, EM_LINEINDEX, (lineNumber - 1), 0&)) + 1
        Case Is = 8 ' = "lineEnd"
        getLineInfo = SendMessageLong(txtObj.hwnd, EM_LINEINDEX, (lineNumber - 1), 0&) + 1 + SendMessageLong(txtObj.hwnd, EM_LINELENGTH, (lineNumber - 1), 0&)
        Case Is = 9 ' = "lineLength"
        getLineInfo = (SendMessageLong(txtObj.hwnd, EM_LINEINDEX, lineNumber, 0&)) + 1 - (SendMessageLong(txtObj.hwnd, EM_LINEINDEX, (lineNumber - 1), 0&)) - 3
    End Select

End Function
Public Function Findfile(xstrfilename) As WIN32_FIND_DATA
    
    Dim Win32Data As WIN32_FIND_DATA
    Dim plngFirstFileHwnd As Long
    Dim plngRtn As Long
    
    plngFirstFileHwnd = FindFirstFile(xstrfilename, Win32Data)  ' Get information of file using API call
    If plngFirstFileHwnd = 0 Then
        Findfile.cFileName = "Error"                            ' If file was not found, return error as name
    Else
        Findfile = Win32Data                                    ' Else return results
    End If
        plngRtn = FindClose(plngFirstFileHwnd)                  ' It is important that you close the handle for FindFirstFile
End Function

Public Sub MakeNormal(hwnd As Long)
    SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub
Public Sub MakeTopMost(hwnd As Long)
    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub

Public Sub SelectionSort(list() As String, ByVal min As Integer, ByVal max As Integer)

Dim i As Integer
Dim j As Integer
Dim best_j As Integer
Dim best_str As String
Dim temp_str As String

    For i = min To max - 1
        best_j = i
        best_str = list(i)
        For j = i + 1 To max
            If StrComp(list(j), best_str, vbTextCompare) < 0 Then
                best_str = list(j)
                best_j = j
            End If
        Next j
        list(best_j) = list(i)
        list(i) = best_str
    Next i
End Sub

