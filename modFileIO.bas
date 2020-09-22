Attribute VB_Name = "modFileIO"
Option Explicit

'Generate temporary Visual Basic files with API
Public Declare Function GetTempFileName Lib "kernel32" _
     Alias "GetTempFileNameA" (ByVal lpszPath As String, _
     ByVal lpPrefixString As String, ByVal wUnique As Long, _
     ByVal lpTempFileName As String) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Function TempFile(strPrefix As String) As String
   
    ' Returns a temporary file name based on the value of strPrefix
    
    Dim strTemp     As String
    Dim lngRet      As Long
    Dim strTempPath As String
    
    On Error GoTo Err_TempFile
    
    strTempPath = Space$(255)
    lngRet = GetTempPath(Len(strTempPath), strTempPath)
    strTemp = Space$(255)
    lngRet = GetTempFileName(strTempPath, strPrefix, 1, ByVal strTemp)
    TempFile = TrimNulls(strTemp)

Exit_TempFile:
    
    On Error GoTo 0
    Exit Function

Err_TempFile:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In modStringStuff, during TempFile" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_TempFile
    End Select
    
End Function
