Attribute VB_Name = "modBuildSQL"
Option Explicit
Public Function BuildADOBottom() As String
Attribute BuildADOBottom.VB_UserMemId = 1610612746
   
    Dim strOutPut As String
    Dim strTemp   As String
    Dim I         As Integer
    
    On Error GoTo Err_BuildADOBottom
    
        If gstrQueryType = "SELECT" Then
            strOutPut = "Set adoRS = New ADODB.Recordset " & vbNewLine
        End If
    strOutPut = strOutPut & "Set adoCN = New ADODB.Connection " & vbNewLine
    strOutPut = strOutPut & "strCN = " & Chr(34) & "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & gstrDBPath & Chr(34) & vbNewLine
    strOutPut = strOutPut & "adoCN.ConnectionString = strCN" & vbNewLine
    strOutPut = strOutPut & "adoCN.Open" & vbNewLine
        If gstrQueryType = "SELECT" Then
            strOutPut = strOutPut & "adoRS.Open(strSQL)" & vbNewLine
            strOutPut = strOutPut & "    If Not adoRS.EOF Then" & vbNewLine
                For I = LBound(gstrRSFields) To UBound(gstrRSFields)
                    strTemp = Trim$(gstrRSFields(I))
                    strTemp = Replace(strTemp, Chr(13), "")
                    strTemp = Replace(strTemp, Chr(10), "")
                    strTemp = Replace(strTemp, vbCrLf, "")
                    strTemp = Replace(strTemp, vbNewLine, "")
                    strTemp = Replace(strTemp, vbTab, "")
                    strOutPut = strOutPut & "        adoRS.Fields(" & Chr(34) & strTemp & Chr(34) & ")" & vbNewLine
                Next
            strOutPut = strOutPut & "        adoRS.MoveNext" & vbNewLine
            strOutPut = strOutPut & "    End If" & vbNewLine
            strOutPut = strOutPut & "    If Not (adoRS Is Nothing) Then" & vbNewLine
            strOutPut = strOutPut & "        adoRS.Close" & vbNewLine
            strOutPut = strOutPut & "        Set adoRS = Nothing" & vbNewLine
            strOutPut = strOutPut & "    End If" & vbNewLine
        Else
            strOutPut = strOutPut & "adoCN.Execute(strSQL)" & vbNewLine
        End If
    strOutPut = strOutPut & "    If Not (adoCN Is Nothing) Then" & vbNewLine
    strOutPut = strOutPut & "        adoCN.Close" & vbNewLine
    strOutPut = strOutPut & "        Set adoCN = Nothing" & vbNewLine
    strOutPut = strOutPut & "    End If" & vbNewLine
    BuildADOBottom = strOutPut

Exit_BuildADOBottom:
    
    On Error GoTo 0
    Exit Function

Err_BuildADOBottom:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In modBuildSQL, during BuildADOBottom" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_BuildADOBottom
    End Select
    
End Function
Public Function BuildADOTop() As String
Attribute BuildADOTop.VB_UserMemId = 1610612745
   
    Dim strOutPut As String

    On Error GoTo Err_BuildADOTop
    
        If gstrQueryType = "SELECT" Then
            strOutPut = "Dim adoRS  As ADODB.Recordset " & vbNewLine
        End If
    strOutPut = strOutPut & "Dim adoCN  As ADODB.Connection " & vbNewLine
    strOutPut = strOutPut & "Dim strCN  As String " & vbNewLine
    strOutPut = strOutPut & "Dim strSQL As String " & vbNewLine
    BuildADOTop = strOutPut

Exit_BuildADOTop:
    
    On Error GoTo 0
    Exit Function

Err_BuildADOTop:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In modBuildSQL, during BuildADOTop" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_BuildADOTop
    End Select

End Function
Public Function BuildDAOBottom() As String
Attribute BuildDAOBottom.VB_UserMemId = 1610612737
  
    Dim strOutPut As String
    Dim strTemp   As String
    Dim I         As Integer
        
    On Error GoTo Err_BuildDAOBottom
    
        If gstrQueryType = "SELECT" Then
            strOutPut = "Set daoRS = New DAO.Recordset " & vbNewLine
        End If
    strOutPut = strOutPut & "Set daoDB = Workspaces(0).OpenDatabase(" & Chr(34) & gstrDBPath & Chr(34) & ")" & vbNewLine
        If gstrQueryType = "SELECT" Then
            strOutPut = strOutPut & "Set daoRS = daoDB.OpenRecordset(strSQL)" & vbNewLine
            strOutPut = strOutPut & "    If Not daoRS.EOF Then" & vbNewLine
                For I = LBound(gstrRSFields) To UBound(gstrRSFields)
                    strTemp = Trim$(gstrRSFields(I))
                    strTemp = Replace(strTemp, Chr(13), "")
                    strTemp = Replace(strTemp, Chr(10), "")
                    strTemp = Replace(strTemp, vbCrLf, "")
                    strTemp = Replace(strTemp, vbNewLine, "")
                    strTemp = Replace(strTemp, vbTab, "")
                    strOutPut = strOutPut & "        daoRS.Fields(" & Chr(34) & strTemp & Chr(34) & ")" & vbNewLine
                Next
            strOutPut = strOutPut & "        daoRS.MoveNext" & vbNewLine
            strOutPut = strOutPut & "    End If" & vbNewLine
            strOutPut = strOutPut & "    If Not (daoRS Is Nothing) Then" & vbNewLine
            strOutPut = strOutPut & "        daoRS.Close" & vbNewLine
            strOutPut = strOutPut & "        Set daoRS = Nothing" & vbNewLine
            strOutPut = strOutPut & "    End If" & vbNewLine
        Else
            strOutPut = strOutPut & "daoDB.Execute(strSQL)" & vbNewLine
        End If
    strOutPut = strOutPut & "    If Not (daoDB Is Nothing) Then" & vbNewLine
    strOutPut = strOutPut & "        Set daoDB = Nothing" & vbNewLine
    strOutPut = strOutPut & "    End If" & vbNewLine
    BuildDAOBottom = strOutPut

Exit_BuildDAOBottom:
    
    On Error GoTo 0
    Exit Function

Err_BuildDAOBottom:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In modBuildSQL, during BuildDAOBottom" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_BuildDAOBottom
    End Select

End Function
Public Function BuildDAOTop() As String
Attribute BuildDAOTop.VB_UserMemId = 1610612736
   
    Dim strOutPut As String

    On Error GoTo Err_BuildDAOTop
    
        If gstrQueryType = "SELECT" Then
            strOutPut = "Dim daoRS  As DAO.Recordset " & vbNewLine
        End If
    strOutPut = strOutPut & "Dim daoDB  As DAO.Database " & vbNewLine
    strOutPut = strOutPut & "Dim strSQL As String " & vbNewLine
    BuildDAOTop = strOutPut

Exit_BuildDAOTop:
    
    On Error GoTo 0
    Exit Function

Err_BuildDAOTop:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In modBuildSQL, during BuildDAOTop" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_BuildDAOTop
    End Select

End Function
Public Function GetStringBetween(strCompleteString As String, strFirst As String, Optional strLast As String, Optional bCaseSensitive As Boolean = False) As String
Attribute GetStringBetween.VB_UserMemId = 1610612742
   
    ' Purpose   : Pass this function a string (strCompleteString),
    '                and it will return a substring consisting of
    '                everything between 2 other specified strings
    '                (ie, everything between strFirst and strLast)
    '             You can also optionally specify if it should be case sensitive (default is False)
    '             Bonus: if you leave out the last string, you'll
    '                      just get the word following the first word
    '                    Or if you leave off the first string, it will start from the
    '                      first character in the main string
    ' Example   : GetStringBetween("Fourscore and seven years ago, our fathers", "and", "our")
    '                would return "seven years ago,"
    ' Parameters: strCompleteString, strFirst, strLast, bCaseSensitive
    ' Returns   : String
    ' Modified  : 3/28/2002 By BB
    
    Dim iPos     As Integer
    Dim iLen     As Integer
    Dim strTemp1 As String
    Dim strTemp2 As String
    
    On Error GoTo Err_GetStringBetween

    ' make sure we have valid values to work with
    If Len(strCompleteString) = 0 Then
        ' no string to parse
        MsgBox "Missing Main String, Nothing to Parse", vbInformation, "Advisory"
        strTemp2 = ""
        GoTo Exit_GetStringBetween
    ElseIf Len(strFirst) = 0 Then
        ' no beginning string, so begin at first character
        iPos = 1
    ElseIf Len(strLast) = 0 Then
        ' no ending string, so we'll make it a space
        strLast = " "
    End If
    ' if no beginning was specified, we can skip this
    If iPos < 1 Then
        ' get the location in the string where our first string occurs
        If bCaseSensitive Then
            ' case sensitive
            iPos = InStr(1, strCompleteString, strFirst, vbBinaryCompare)
        Else
            ' case insensitive
            iPos = InStr(1, strCompleteString, strFirst, vbTextCompare)  ' default
        End If
    End If
    ' assuming we did find the first string...
    If iPos > 0 Then
        ' extract everything to the right of the first string;
        ' we use the expression
        '     Len(strCompleteString) - (iPos + Len(Trim$(strFirst)
        ' to determine where the first string actually ends,
        ' the Trim$ call makes sure we don't include any spaces the user may have passed in
        ' (you have to pass in the spaces around a word to distinguish a complete word
        ' from a string that may appear within a word, eg, the "and" in "thousand" would
        ' mess us up if we had called it like this:
        '    GetStringBetween("Four thousand and seven years ago", "and", "ago")
        ' so the right way to call it would be this:
        '    GetStringBetween("Four thousand and seven years ago", " and ", "ago")
        '
        ' I hope that makes it clear!
        If iPos = 1 Then
            strTemp1 = Trim$(Right$(strCompleteString, Len(strCompleteString)))
        Else
            strTemp1 = Trim$(Right$(strCompleteString, Len(strCompleteString) - (iPos + Len(Trim$(strFirst)))))
        End If
    End If
    If (LCase$(strFirst) = " inner join ") And (LCase$(strLast) = " on ") Then
        iLen = Len(strTemp1)
        If bCaseSensitive Then
            ' case sensitive
            iPos = InStrRev(strTemp1, strLast, iLen, vbBinaryCompare)
        Else
            ' case insensitive
            iPos = InStrRev(strTemp1, strLast, iLen, vbTextCompare)  ' default
        End If
        If iPos > 0 Then
            strTemp2 = " INNER JOIN " & Trim$(Left$(strTemp1, iPos - 1)) & " ON "
        Else
            strTemp2 = strTemp1
        End If
    Else
        If bCaseSensitive Then
            ' case sensitive
            iPos = InStr(1, strTemp1, strLast, vbBinaryCompare)
        Else
            ' case insensitive
            iPos = InStr(1, strTemp1, strLast, vbTextCompare)  ' default
        End If
        If iPos > 0 Then
            strTemp2 = Trim$(Left$(strTemp1, iPos - 1))
        Else
            strTemp2 = strTemp1
        End If
    End If
    
Exit_GetStringBetween:
    
    On Error Resume Next
    GetStringBetween = strTemp2
    On Error GoTo 0
    Exit Function

Err_GetStringBetween:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In modBuildSQL, during GetStringBetween" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            strTemp2 = ""
            Resume Exit_GetStringBetween
    End Select

End Function
