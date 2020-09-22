Attribute VB_Name = "modGetDatabase"
Option Explicit
Public Sub GetDatabase()
Attribute GetDatabase.VB_UserMemId = 1610612736
   
    Dim cnn        As ADODB.Connection
    Dim rst        As ADODB.Recordset
    Dim strSQL     As String
    Dim strConnect As String
    
    On Error GoTo Err_GetDatabase
    
    Set cnn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    strConnect = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & frmSelectData.txtDBPath.Text
    cnn.Open strConnect

Exit_GetDatabase:
    
    On Error Resume Next
        If Not (cnn Is Nothing) Then
            Set cnn = Nothing
        End If
        If Not (rst Is Nothing) Then
            Set rst = Nothing
        End If
    On Error GoTo 0
    Exit Sub

Err_GetDatabase:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In modGetDatabase, during GetDatabase" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_GetDatabase
    End Select

End Sub
