Attribute VB_Name = "modGlobalSearch"
' module for the intextsearch function
Public Function GetSelectedFile(strPath As String) As String

If Right(strPath, 1) <> "\" Then
 GetSelectedFile = strPath & "\" & frmGlobal.File1.FileName
  Else
   GetSelectedFile = strPath & frmGlobal.File1.FileName
    
End If
    
End Function
