VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'*********************************************************************
'       FILESEARCH CLASS MODULE
'
' Search Files On a Local Drive Very Easy
' With This Class Module.
' Save This File on Your Hardrive. When
' Needed In a Project, Just Pick "Add Class Module"
' In Project-Menu, Choose "Existing", Browse To This
' File. Following Shows How To Use It On a Form.
' OfCourse You Need a CommandButton Named cmdSearch,
' 2 TextBoxes named txtPath (Start Directory) and txtFilter,
' Label1 and Label2 to Show Info.
'
' FileSearch Class Was Writen By $$$$$
' (c) FiveBucksEntertainment
' Contact/BugReports to fivebucks@swipnet.se
'
' Good Luck !!!
'
' ///  COPY --> PASTE ///
'
'Private Sub cmdSearch_Click()
'Dim fs1 As New FileSearch      ' Create a New Instans of Class
'Dim iSize as Long              ' Filesize
'Dim i as Long
'Dim X as String
'
'    On Error GoTo searcherror:
'
'    With fs1
'        .Reset                          ' Reset, only needed if class is global
'        .ProgressLabel = Label1         ' if tou want to show progress info on Label1
'        .StartSearchHere = txtPath      ' required, start here
'        .Filters = txtFilter            ' Filters ex. "*.bmp;*.exe;explorer.exe"
'        .ListDirs = False               ' boolean
'        .SearchInSubDirs = True         ' = check1.value works fine !!!
'        .Start                          ' GO !!!!
'
'       ' Loop through found items
'        For i = 0 To .FileCount - 1
'
'            .ListIndex = i
'
'            'Do What you want with search result !!!
'                'If .IsDirectory Then         ' if directory...
'                'X = .FileName                ' get filename
'                'iSize = (.FileLenght / 1024) + 0.5   ' calculate filesize in Kb
'                'X = iSize & " Kb"
'                'X = .FileAttrib              ' get file attributes
'                'X = fs1.Location             ' get file locations , full path
'
'        Next i
'
'        'Label2 = "Found: " & .FileCount & " Files"
'
'    End With
'
'Exit Sub
'searcherror:
'    MsgBox Err.Description, vbExclamation, "Runtime Error: " & Err.Number
'    fs1.Reset
'End Sub
' ///// END OF CODE ////
'
Option Explicit

Public SearchInSubDirs As Boolean
Public StartSearchHere As String
Public Filters As String
Public ListDirs As Boolean


' local copies
Dim ShowProgress As Boolean
Dim fslFileCount As Long
Dim fslFileLenght As Long
Dim fslFileName As String
Dim fslFilePath As String
Dim fslFileAttrib As String
Dim fslIsDir As Boolean
Dim fslListIndex As Long
Dim fslLabel As Control

' Local Arrays
Dim aFiles() As String
Dim aPath() As String
Dim aFileLen() As Long
Dim aAttrib() As Long
Dim aFilter() As String     ' dynamic array containing all filterparameters

Dim NumOfFilters As Integer

Public Enum FileSearchConstants
    fsScanSubs = 1
    fsNoSubs = 0
    fsListAll = 23
    fsNormal = 0
    fsHiddenFiles = 2
    fsAllFiles = 6
End Enum

Public Sub Reset()
    ReDim aFiles(1)
    ReDim aPath(1)
    ReDim aFileLen(1)
    ReDim aAttrib(1)
    ReDim aFilter(1)
    
    StartSearchHere = ""
    Filters = "*"
    SearchInSubDirs = True
    'SearchMode = fsAllFiles
    ListDirs = False
    fslFileCount = 0
    ShowProgress = False
    Set fslLabel = Nothing
End Sub
Public Property Let ListIndex(iValue As Long)
Dim iFileAttrib As Long

    If iValue > fslFileCount - 1 Then
        Err.Raise 10002, , "Can't Assign to Index = " & iValue & vbCrLf & "Value out of range"
    ElseIf iValue < 0 Then
        Err.Raise 10003, , "Index must be a positive value." & vbCrLf & "Value out of range"
    Else
        
        fslListIndex = iValue
        
        fslFileName = aFiles(iValue)
        fslFilePath = aPath(iValue)

        fslFileLenght = aFileLen(iValue)
        iFileAttrib = aAttrib(iValue)
        
        fslIsDir = False
        
        fslFileAttrib = ""
        If iFileAttrib >= 32 Then fslFileAttrib = "A": iFileAttrib = iFileAttrib - 32
        If iFileAttrib >= 16 Then fslIsDir = True: iFileAttrib = iFileAttrib - 16
        If iFileAttrib >= 4 Then fslFileAttrib = fslFileAttrib & "S": iFileAttrib = iFileAttrib - 4
        If iFileAttrib >= 2 Then fslFileAttrib = fslFileAttrib & "H": iFileAttrib = iFileAttrib - 2
        If iFileAttrib >= 1 Then fslFileAttrib = fslFileAttrib & "R"
    
    End If
        
End Property
Public Property Let ProgressLabel(lblLabel As Control)
    Set fslLabel = lblLabel
    ShowProgress = True
End Property
Public Property Get ListIndex() As Long
    ListIndex = fslListIndex
End Property
Public Property Get FileName() As String
    FileName = fslFileName
End Property
Public Property Get FileLenght() As Long
    FileLenght = fslFileLenght
End Property
Public Property Get FileCount() As Long
    FileCount = fslFileCount
End Property
Public Property Get FileAttrib() As String
    FileAttrib = fslFileAttrib
End Property
Public Property Get IsDirectory() As Boolean
   IsDirectory = fslIsDir
End Property
Public Property Get Location() As String
 Location = fslFilePath
End Property

Public Sub Start()
Dim counter%, index%

    CancelDirSearch = False
    
    If Right$(StartSearchHere, 1) <> "\" Then StartSearchHere = StartSearchHere & "\"
    If Dir(StartSearchHere, vbDirectory) = "" Then Err.Raise 10001, , "The Path Doesn't exist.": Exit Sub
    
    ' separate all filter parameters
    counter = 1: index = 1       ' reset start value
    
    ReDim aFilter(5)    ' reallocate
    
    Filters = LCase(Filters)    ' convert to lowercase
    
    Do While True
        ' check if ";" is found
        If Mid(Filters, counter, 1) = ";" Then
            
            aFilter(NumOfFilters) = Mid(Filters, index, counter - index)
            
            index = counter + 1
            NumOfFilters = NumOfFilters + 1
            
            ' reallocate array "aFilter" if needed
            If NumOfFilters >= UBound(aFilter) Then
                ReDim Preserve aFilter(NumOfFilters + 3)
            End If
            
        End If
        counter = counter + 1
        If counter >= Len(Filters) Then Exit Do
        If CancelDirSearch = True Then Exit Sub
    Loop
    aFilter(NumOfFilters) = Right(Filters, Len(Filters) - index + 1)
    ' separate filter parameters DONE!!!
    
    ' Start Search...
    GetFilesNow StartSearchHere
    
    If ShowProgress Then fslLabel.Caption = "Listing Files..."
    DoEvents
End Sub
Private Sub GetFilesNow(ByVal RootDir As String)
Dim nSubDir As Long         ' räknar subdirs
Dim sSubDirs() As String    ' dynamisk array som lagrar underkataloger
Dim sTemp As String
Dim i&, counter&
Dim index As Integer
Dim bIsDir As Boolean

    nSubDir = -1
    
    ReDim sSubDirs(10)  'ökar arrayn till 10 items
    
    If ShowProgress Then fslLabel.Caption = "Searching: " & RootDir
    
    DoEvents
    
    sTemp = Dir$(RootDir, 23)
    
    Do While Len(sTemp) > 0
        DoEvents
        If sTemp <> "." And sTemp <> ".." Then
            bIsDir = False
            ' if directory
            If (GetAttr(RootDir & sTemp) And vbDirectory) = vbDirectory Then
                If ListDirs Then ListFile RootDir, sTemp
                nSubDir = nSubDir + 1
                
                If nSubDir = UBound(sSubDirs) Then
                    ReDim Preserve sSubDirs(nSubDir + 10)
                End If
                
                sSubDirs(nSubDir) = RootDir & sTemp & "\"
                bIsDir = True
            Else
                ListFile RootDir, sTemp
            End If
            
        End If
    sTemp = Dir$
    If CancelDirSearch = True Then Exit Sub
    Loop
        
    If SearchInSubDirs Then
        For i = 0 To UBound(sSubDirs)
            If sSubDirs(i) = "" Then Exit For
            GetFilesNow sSubDirs(i)
        Next i
    End If
End Sub
Private Sub ListFile(spath As String, sfile As String)
Dim index As Integer

    For index = 0 To NumOfFilters
        
        If LCase(sfile) Like aFilter(index) Then
                
            aFiles(fslFileCount) = sfile
            aPath(fslFileCount) = spath
            aFileLen(fslFileCount) = FileLen(spath & sfile)
            aAttrib(fslFileCount) = GetAttr(spath & sfile)
            
            'Debug.Print aFiles(FileCount) & "   " & aAttrib(fslFileCount)
            fslFileCount = fslFileCount + 1
            
            If fslFileCount = UBound(aFiles) Then
                ReDim Preserve aFiles(fslFileCount + 100)
                ReDim Preserve aPath(fslFileCount + 100)
                ReDim Preserve aFileLen(fslFileCount + 100)
                ReDim Preserve aAttrib(fslFileCount + 100)
            End If
            Exit Sub ' files already listed so why continue...
        End If
        If CancelDirSearch = True Then Exit Sub
    Next index

End Sub
Private Sub Class_Terminate()
    
    'MsgBox "terminating..."
    ReDim aFiles(0)
    ReDim aPath(0)
    ReDim aFileLen(0)
    ReDim aAttrib(0)
    ReDim aFilter(0)
    
End Sub
Private Sub Class_Initialize()
    'MsgBox "In Constructor"
    ReDim aFiles(1)
    ReDim aPath(1)
    ReDim aFileLen(1)
    ReDim aAttrib(1)
    ReDim aFilter(1)
    
    StartSearchHere = ""
    Filters = "*"
    SearchInSubDirs = True
    'SearchMode = fsAllFiles
    ListDirs = False
    fslFileCount = 0
    ShowProgress = False
    Set fslLabel = Nothing
End Sub

