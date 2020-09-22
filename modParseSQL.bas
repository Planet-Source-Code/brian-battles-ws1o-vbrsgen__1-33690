Attribute VB_Name = "modParseSQL"
Option Explicit

Public gbAlreadyOpen  As Boolean
Public gstrTableName  As String
Public gstrDBPath     As String
Public gstrRSFields() As String
Public gstrQueryType  As String
Public Function BuildSQL(strSQLStatement As String) As String
Attribute BuildSQL.VB_UserMemId = 1610612736
   
    ' Purpose   : generate SQL strings to be pasted into VB code
    ' Parameters: strSQLStatement
    ' Returns   : String
    ' Modified  : 3/27/2002 By BB
    
    Dim strFirst       As String
    Dim strAllButFirst As String
    Dim strFieldList() As String
    Dim strValueList() As String
    Dim strOutPut      As String
    Dim iPos           As Integer
        
    On Error GoTo Err_BuildSQL
    
    strSQLStatement = Replace(strSQLStatement, Chr(13), " ")
    strSQLStatement = Replace(strSQLStatement, Chr(10), " ")
    strSQLStatement = Replace(strSQLStatement, "  ", " ")
    strFirst = LCase$(Left$(strSQLStatement, 6))
    Select Case strFirst
        Case "select"
            gstrQueryType = "SELECT"
            If LCase$(Left$(strSQLStatement, 11)) = "select into" Then
                ' handle select into statement
                iPos = 11
                strAllButFirst = Right$(strSQLStatement, Len(strSQLStatement) - iPos)
            Else
                strOutPut = GetSelect(strSQLStatement)
                GoTo Exit_BuildSQL
            End If
        Case "insert"
            gstrQueryType = "INSERT"
            strOutPut = GetInsertInto(strSQLStatement)
        Case "update"
            gstrQueryType = "UPDATE"
            strOutPut = GetUpdate(strSQLStatement)
        Case "delete"
            gstrQueryType = "DELETE"
            iPos = 6
            strAllButFirst = Right$(strSQLStatement, Len(strSQLStatement) - iPos)
            ' delete from TABLENAME where FIELD1 = VALUE1 AND/OR FIELD2 = VALUE2, etc
        Case Else
            ' huh?
    End Select
    strOutPut = strOutPut & HandleWhere(strSQLStatement)

Exit_BuildSQL:
    
    On Error Resume Next
    BuildSQL = strOutPut
    Erase strFieldList
    Erase strValueList
    On Error GoTo 0
    Exit Function

Err_BuildSQL:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In modBuildSQL, during BuildSQL" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            strOutPut = ""
            Resume Exit_BuildSQL
    End Select

End Function
Public Function CleanUpInnerJoins(strSQLStatement As String) As String
Attribute CleanUpInnerJoins.VB_UserMemId = 1610612737
   
    ' need to translate clauses like
    'INNER JOIN (tblLocations INNER JOIN tblUsers ON tblLocations.LocationCode=tblUsers.LocationCode) ON
    ' to something like
    ' FROM tblLocations, tblUsers WHERE tblLocations.LocationCode = tblUsers.LocationCode
    
    Dim astrWords       As String
    Dim strWords()      As String
    Dim strSentenceChar As String
    Dim rastrWords()    As String
    Dim strTables       As String
    Dim strJoinConds    As String
    Dim strAnd          As String
    Dim strOr           As String
    Dim strTemp         As String
    Dim lngNoOfWords    As Long
    Dim lngLCV          As Long
    Dim lngRowNo        As Long
    Dim I               As Integer
    Dim bGotOneInner    As Boolean
    Dim bGotOneJoin     As Boolean
    Dim bGotOneOn       As Boolean
    Dim bGotOneOr       As Boolean
    Dim bGotOneAnd      As Boolean
    
    On Error GoTo Err_CleanUpInnerJoins
    
    strSQLStatement = Replace(strSQLStatement, vbCrLf, " ")
    strSQLStatement = Replace(strSQLStatement, vbTab, " ")
    strSQLStatement = Replace(strSQLStatement, Chr(10), " ")
    strSQLStatement = Replace(strSQLStatement, Chr(13), " ")
    'count the number of words. Number of words = number of spaces plus one
    For lngLCV = 1 To Len(strSQLStatement)
        If Mid$(strSQLStatement, lngLCV, 1) = Space$(1) Then
            lngNoOfWords = lngNoOfWords + 1
        End If
    Next lngLCV
    'make the array big enough to hold the words
    ReDim rastrWords(lngNoOfWords + 1)
    'put each word into a row in the array
    For lngLCV = 1 To Len(strSQLStatement)
        strSentenceChar = Mid$(strSQLStatement, lngLCV, 1)
            If strSentenceChar <> Space$(1) Then
                rastrWords(lngRowNo) = rastrWords(lngRowNo) & strSentenceChar
            Else
                lngRowNo = lngRowNo + 1
            End If
    Next lngLCV
    For I = LBound(rastrWords) To UBound(rastrWords)
        Select Case LCase$(rastrWords(I))
            Case "and"
                bGotOneAnd = True
            Case "or"
                bGotOneOr = True
            Case "inner"
                bGotOneInner = True
            Case "join"
                bGotOneJoin = True
            Case "on"
                bGotOneOn = True
            Case Else
                If bGotOneOn = True Then
                    If strJoinConds = "" Then
                        strJoinConds = rastrWords(I)
                    Else
                        strJoinConds = strJoinConds & ", " & rastrWords(I)
                    End If
                    bGotOneOn = False
                    GoTo NextOne
                End If
                If (bGotOneInner = True) And (bGotOneJoin = True) Then
                    If strTables = "" Then
                        strTables = rastrWords(I)
                    Else
                        strTables = strTables & ", " & rastrWords(I)
                    End If
                    bGotOneInner = False
                    bGotOneJoin = False
                    GoTo NextOne
                End If
                If bGotOneAnd = True Then
                    If strAnd = "" Then
                        strAnd = rastrWords(I)
                    Else
                        If InStr(strAnd, rastrWords(I)) Then
                            ' don't get a duplicate
                        Else
                            strAnd = strAnd & " AND " & rastrWords(I)
                        End If
                    End If
                    bGotOneAnd = False
                    GoTo NextOne
                End If
                If bGotOneOr = True Then
                    If strOr = "" Then
                        strOr = rastrWords(I)
                    Else
                        If InStr(strOr, rastrWords(I)) Then
                            ' don't get a duplicate
                        Else
                            strOr = strOr & " OR " & rastrWords(I)
                        End If
                    End If
                    bGotOneOr = False
                    GoTo NextOne
                End If
        End Select

NextOne:

    Next
    ' clean up the resulting strings
    strTables = Replace(strTables, "(", "")
    strTables = Replace(strTables, ")", "")
    strJoinConds = Replace(strJoinConds, ")", "")
    strJoinConds = Replace(strJoinConds, "(", "")
    strJoinConds = Replace(strJoinConds, "=", " = ")
    strJoinConds = Replace(strJoinConds, "> =", " >= ")
    strJoinConds = Replace(strJoinConds, "= <", " =< ")
    strJoinConds = Replace(strJoinConds, "< >", " <> ")
    strJoinConds = Replace(strJoinConds, ";", "")
    strJoinConds = Replace(strJoinConds, ",", " AND ")
    strOr = Replace(strOr, ";", "")
    strOr = Replace(strOr, ")", "")
    strOr = Replace(strOr, "(", "")
    strAnd = Replace(strAnd, ";", "")
    strAnd = Replace(strAnd, ")", "")
    strAnd = Replace(strAnd, "(", "")
        If Trim$(strJoinConds) = "" Then
            strJoinConds = ""
        Else
            strJoinConds = " WHERE " & strJoinConds
        End If
        If Trim$(strAnd) = "" Then
            strAnd = ""
        Else
            If Left$(strJoinConds, 6) = " WHERE" Then
                strAnd = " AND " & strAnd
            Else
                strAnd = " WHERE " & strAnd
            End If
        End If
        If Trim$(strOr) = "" Then
            strOr = strOr
        Else
            If Left$(strJoinConds, 6) = " WHERE" Then
                If Trim$(strAnd) = "" Then
                    strOr = " "
                Else
                    strOr = " WHERE " & strOr
                End If
            Else
                strOr = " OR " & strOr
            End If
        End If
    strTemp = strTables & strJoinConds & strAnd & strOr
    strTemp = Replace(strTemp, " = ", "=")
    strTemp = Replace(strTemp, "> =", ">=")
    strTemp = Replace(strTemp, "< =", "<=")
    strTemp = Replace(strTemp, "  ", " ")
    CleanUpInnerJoins = strTemp

Exit_CleanUpInnerJoins:
    
    On Error Resume Next
    Erase strWords
    Erase rastrWords
    On Error GoTo 0
    Exit Function

Err_CleanUpInnerJoins:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In modBuildSQL, during CleanUpInnerJoins" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_CleanUpInnerJoins
    End Select

End Function
Public Function GetDelete(strTheSQLStatement As String) As String
Attribute GetDelete.VB_UserMemId = 1610612738
    
    Dim strAllButFirst  As String
    Dim strBeginning    As String
    Dim strTableName    As String
    Dim strFields       As String
    Dim strValues       As String
    Dim strFieldPart()  As String
    Dim strFieldList()  As String
    Dim strValuePart()  As String
    Dim strValueList()  As String
    Dim strOutPut       As String
    Dim strTemp         As String
    Dim iPos            As Integer
    Dim iSpacePos       As Integer
    Dim I               As Integer
    
    On Error GoTo Err_GetDelete
    
    strAllButFirst = Trim$(Right$(strTheSQLStatement, Len(strTheSQLStatement) - 6))
    ' Delete * FROM TABLENAME Where FIELD1 = VALUE1 AND/OR FIELD2 = VALUE2, etc
    ' extract TABLE NAME
    iSpacePos = InStr(strAllButFirst, " ")
    strTableName = Trim$(Left$(strAllButFirst, iSpacePos))
    ' DELETE is 6 characters long, so...
    strAllButFirst = Trim$(Right$(strTheSQLStatement, Len(strTheSQLStatement) - 6))
    strBeginning = UCase$(Trim$(Left$(strTheSQLStatement, 6)))
    ' extract TABLE NAME
    iSpacePos = InStr(strAllButFirst, " ")
    strTableName = Left$(strAllButFirst, iSpacePos)
    strAllButFirst = Replace(strAllButFirst, strTableName, "")
    ' now let's see if we have a WHERE Clause...
    iPos = InStr(1, strAllButFirst, " where ", vbTextCompare)
        If iPos > 0 Then
            ' looks like we do have a WHERE Clause...
        Else
            ' no WHERE, just a straight delete-eveything statement
            strAllButFirst = Trim$(Replace(strAllButFirst, " set ", " SET ", vbTextCompare))
            iPos = InStr(1, strAllButFirst, "SET ", vbTextCompare)
            strFields = Trim$(Right$(strAllButFirst, Len(strAllButFirst) - 4))
            strFields = Replace(strFields, "(", "")
            strFields = Replace(strFields, ")", "")
            strFieldPart = Split(strFields, ",")
            ReDim strFieldList(UBound(strFieldPart))
                For I = LBound(strFieldPart) To UBound(strFieldPart)
                    iPos = InStr(strFieldPart(I), "=")
                        If iPos > 0 Then
                            strTemp = Trim$(Left$(strFieldPart(I), iPos - 1))
                            strFieldList(I) = strTemp
                        Else
                            strFieldList = Split(strFieldPart(I), "=")
                        End If
                Next
            strValues = Trim$(Right$(strAllButFirst, Len(strAllButFirst) - 4))
            strValues = Replace(strValues, "(", "")
            strValues = Replace(strValues, ")", "")
            strValuePart = Split(strValues, ",")
            ReDim strValueList(UBound(strValuePart))
                For I = LBound(strValuePart) To UBound(strValuePart)
                    iPos = InStr(strFieldPart(I), "=")
                        If iPos > 0 Then
                            strTemp = Trim$(Right$(strValuePart(I), iPos - 1))
                            strValueList(I) = strTemp
                        Else
                            strValueList = Split(strValuePart(I), "=")
                        End If
                Next
            strOutPut = ""
            strOutPut = "strSQL = " & Chr(34) & strBeginning & " " & Chr(34) & vbNewLine
            strOutPut = strOutPut & "strSQL = strSQL & " & Chr(34) & strTableName & " " & Chr(34) & vbNewLine
            strOutPut = strOutPut & "strSQL = strSQL & " & Chr(34) & "SET " & Chr(34) & vbNewLine
            strOutPut = strOutPut & "strSQL = strSQL & " & Chr(34) & "(" & Chr(34) & vbNewLine
                For I = LBound(strFieldList) To UBound(strFieldList)
                        If I = UBound(strFieldList) Then
                            strOutPut = strOutPut & "strSQL = strSQL & " & Chr(34) & Trim$(strFieldList(I)) & " = " & Trim$(strValueList(I)) & Chr(34) & vbNewLine
                        Else
                            strOutPut = strOutPut & "strSQL = strSQL & " & Chr(34) & Trim$(strFieldList(I)) & " = " & Trim$(strValueList(I)) & ", " & Chr(34) & vbNewLine
                        End If
                Next
            strOutPut = strOutPut & "strSQL = strSQL & " & Chr(34) & ")" & Chr(34) & vbNewLine
            GetDelete = strOutPut
        End If
        
Exit_GetDelete:

    On Error Resume Next
    Erase strFieldList
    Erase strValueList
    Erase strFieldPart
    Erase strValuePart
    On Error GoTo 0
    Exit Function

Err_GetDelete:
        
    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In modBuildSQL, during GetDelete" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_GetDelete
    End Select

End Function
Private Function GetInsertInto(strTheSQLStatement As String) As String
   
    ' Purpose   :
    ' Parameters: strTheSQLStatement
    ' Returns   : String
    ' Modified  : 3/27/2002 By BB

    Dim strAllButFirst As String
    Dim strBeginning   As String
    Dim strTableName   As String
    Dim strFields      As String
    Dim strValues      As String
    Dim strFieldList() As String
    Dim strValueList() As String
    Dim strOutPut      As String
    Dim strValOrSel    As String
    Dim strEndFrom     As String
    Dim strTemp        As String
    Dim iPos           As Integer
    Dim iSpacePos      As Integer
    Dim I              As Integer

    On Error GoTo Err_GetInsertInto
    
    strAllButFirst = Trim$(Right$(strTheSQLStatement, Len(strTheSQLStatement) - 11))
    ' get rid of any WHERE clause...
    iPos = InStr(1, strAllButFirst, " where ", vbTextCompare)
        If iPos > 0 Then
            strAllButFirst = Trim$(Left$(strAllButFirst, iPos - 1))
        End If
    strBeginning = UCase$(Trim$(Left$(strTheSQLStatement, 12)))
    ' extract TABLE NAME
    iSpacePos = InStr(strAllButFirst, " ")
    strTableName = Left$(strAllButFirst, iSpacePos)
    strAllButFirst = Replace(strAllButFirst, strTableName, "")
    strAllButFirst = Replace(strAllButFirst, "values", " VALUES ", vbTextCompare)
    iPos = InStr(1, strAllButFirst, "values ", vbTextCompare)
        If iPos < 1 Then
            iPos = InStr(1, strAllButFirst, "select ", vbTextCompare)
            strValOrSel = "SELECT "
        Else
            strValOrSel = " VALUES "
        End If
    strFields = Trim$(Left$(strAllButFirst, iPos - 1))
    strFields = Replace(strFields, "(", "")
    strFields = Replace(strFields, ")", "")
    strFields = Trim$(strFields)
    strFieldList = Split(strFields, ",")
    strValues = Trim$(Right$(strAllButFirst, iPos + 8)) ' have to allow for invisible CRs and LFs, I guess
    strValues = Replace(strValues, "(", "")
    strValues = Replace(strValues, ")", "")
    If InStr(1, strValues, "from ", vbTextCompare) Then
        strEndFrom = Right$(strValues, Len(strValues) - (InStr(1, strValues, "from ", vbTextCompare) - 1))
        strValues = Replace(strValues, strEndFrom, "")
        strEndFrom = Replace(strEndFrom, Chr(13), "")
        strEndFrom = Replace(strEndFrom, Chr(10), "")
        strEndFrom = Replace(strEndFrom, ";", "")
        strValues = Trim$(strValues)
        strValues = Replace(strValues, Chr(13), "")
        strValues = Replace(strValues, Chr(10), "")
    End If
    strValueList = Split(strValues, ",")
    strOutPut = ""
    strOutPut = "strSQL = " & Chr(34) & strBeginning & " " & Chr(34) & vbNewLine
    strOutPut = strOutPut & "strSQL = strSQL & " & Chr(34) & strTableName & " " & Chr(34) & vbNewLine
    strOutPut = strOutPut & "strSQL = strSQL & " & Chr(34) & "(" & Chr(34) & vbNewLine
    gstrRSFields = strFieldList
        For I = LBound(strFieldList) To UBound(strFieldList)
            ' clean out all CRs and LFs and such
            strTemp = Replace(strFieldList(I), Chr(13), "")
            strTemp = Replace(strTemp, Chr(10), "")
            strTemp = Replace(strTemp, vbCrLf, "")
            strTemp = Replace(strTemp, vbNewLine, "")
            strTemp = Replace(strTemp, vbTab, "")
            strTemp = Trim$(strTemp)
            If I = UBound(strFieldList) Then
                strOutPut = strOutPut & "strSQL = strSQL & " & Chr(34) & strTemp & Chr(34) & vbNewLine
            Else
                strOutPut = strOutPut & "strSQL = strSQL & " & Chr(34) & strTemp & ", " & Chr(34) & vbNewLine
            End If
        Next
    strOutPut = strOutPut & "strSQL = strSQL & " & Chr(34) & ")" & " " & Chr(34) & vbNewLine
        If InStr(1, strValOrSel, "values", vbTextCompare) Then
            strOutPut = strOutPut & "strSQL = strSQL & " & Chr(34) & strValOrSel & " " & Chr(34) & vbNewLine
            strOutPut = strOutPut & "strSQL = strSQL & " & Chr(34) & "(" & Chr(34) & vbNewLine
        Else
            strOutPut = strOutPut & "strSQL = strSQL & " & Chr(34) & "(" & Chr(34) & vbNewLine
            strOutPut = strOutPut & "strSQL = strSQL & " & Chr(34) & strValOrSel & " " & Chr(34) & vbNewLine
        End If
        For I = LBound(strValueList) To UBound(strValueList)
            ' clean out all CRs and LFs and such
            strTemp = Replace(strValueList(I), Chr(13), "")
            strTemp = Replace(strTemp, Chr(10), "")
            strTemp = Replace(strTemp, vbCrLf, "")
            strTemp = Replace(strTemp, vbNewLine, "")
            strTemp = Replace(strTemp, vbTab, "")
            strTemp = Trim$(strTemp)
            If I = UBound(strValueList) Then
                strOutPut = strOutPut & "strSQL = strSQL & " & Chr(34) & strTemp & Chr(34) & vbNewLine
            Else
                strOutPut = strOutPut & "strSQL = strSQL & " & Chr(34) & strTemp & ", " & Chr(34) & vbNewLine
            End If
        Next
    strOutPut = strOutPut & "strSQL = strSQL & " & Chr(34) & strEndFrom & Chr(34) & vbNewLine
    strOutPut = strOutPut & "strSQL = strSQL & " & Chr(34) & ")" & Chr(34)
    strOutPut = Replace(strOutPut, ";", "")
    GetInsertInto = strOutPut

Exit_GetInsertInto:
    
    On Error Resume Next
    Erase strFieldList
    Erase strValueList
    On Error GoTo 0
    Exit Function

Err_GetInsertInto:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In modBuildSQL, during GetInsertInto" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_GetInsertInto
    End Select

End Function
Public Function GetSelect(strSQLSelect As String) As String
Attribute GetSelect.VB_UserMemId = 1610612740
   
    Dim strAllButFirst As String
    Dim strTableNames  As String
    Dim strTables()    As String
    Dim strFieldList() As String
    Dim strFieldNames  As String
    Dim strValuePart() As String
    Dim strTemp        As String
    Dim strOutPut      As String
    Dim strDistinct    As String
    Dim strIJ          As String
    Dim strWhere       As String
    Dim iPos           As Integer
    Dim iSpacePos      As Integer
    Dim I              As Integer
    
    On Error GoTo Err_GetSelect
    
    ' SELECT is 6 characters long, so...
    strAllButFirst = Trim$(Right$(strSQLSelect, Len(strSQLSelect) - 6))
        ' is DISTINCT in there?
        If InStr(1, strAllButFirst, "distinct ", vbTextCompare) Then
            strDistinct = "DISTINCT "
            strAllButFirst = Trim$(Replace(strAllButFirst, "distinctrow ", "", , vbTextCompare))
        ElseIf InStr(1, strAllButFirst, "distinctrow ", vbTextCompare) Then
            strDistinct = "DISTINCTROW "
            strAllButFirst = Trim$(Replace(strAllButFirst, "distinctrow ", "", , vbTextCompare))
        End If
    ' get any WHERE clause...
    iPos = InStr(1, strAllButFirst, " where ", vbTextCompare)
        If iPos > 0 Then
            strAllButFirst = Trim$(Left$(strAllButFirst, iPos - 1))
        End If
    iPos = InStr(1, strSQLSelect, " where ", vbTextCompare)
    strWhere = Trim$(Right$(strSQLSelect, Len(strSQLSelect) - iPos))
    ' extract FIELD NAME(S)
    iSpacePos = InStr(strAllButFirst, "FROM ")
    strFieldNames = Trim$(Left$(strAllButFirst, iSpacePos - 1))
    strAllButFirst = Trim$(Right$(strAllButFirst, Len(strAllButFirst) - (iSpacePos + 4)))
    ' extract TABLE NAME(S) and such from the Inner Join clauses
    strIJ = CleanUpInnerJoins(strAllButFirst)
        If InStr(1, strAllButFirst, " where ", vbTextCompare) Then
            If InStr(1, strAllButFirst, "from ", vbTextCompare) Then
                strTableNames = GetStringBetween(strAllButFirst, "from ", " where ")
            Else
                strTableNames = GetStringBetween(strAllButFirst, "", " where ")
            End If
        Else
            If InStr(1, strAllButFirst, " inner ", vbTextCompare) Then
                strTableNames = GetStringBetween(strAllButFirst, "", " inner ")
            End If
        End If
        If Trim$(strTableNames) = "" Then
            strTableNames = Trim$(strAllButFirst)
        End If
    strIJ = Replace(strIJ, Chr(34), Chr(39))
    If Trim$(strIJ) = "" Then
        strIJ = HandleWhere(strSQLSelect)
        strIJ = SplitTheWhere(strWhere)
        Debug.Print "strIJ: " & strIJ
        strValuePart = Split(strIJ, ",")
    Else
        strValuePart = Split(strIJ, " ")
    End If
    strTables = Split(strTableNames, ",")
    strOutPut = ""
    strOutPut = "strSQL = " & Chr(34) & "SELECT " & strDistinct & Chr(34) & vbNewLine
    strFieldList = Split(strFieldNames, ",")
    gstrRSFields = strFieldList
        For I = LBound(strFieldList) To UBound(strFieldList)
            ' clean out all CRs and LFs and such
            strTemp = Replace(strFieldList(I), Chr(13), "")
            strTemp = Replace(strTemp, Chr(10), "")
            strTemp = Replace(strTemp, vbCrLf, "")
            strTemp = Replace(strTemp, vbNewLine, "")
            strTemp = Replace(strTemp, vbTab, "")
            strTemp = Trim$(strTemp)
                If I = UBound(strFieldList) Then
                    strOutPut = strOutPut & "strSQL = strSQL & " & Chr(34) & strTemp & " " & Chr(34) & vbNewLine
                Else
                    strOutPut = strOutPut & "strSQL = strSQL & " & Chr(34) & strTemp & ", " & Chr(34) & vbNewLine
                End If
        Next
    strOutPut = strOutPut & "strSQL = strSQL & " & Chr(34) & "FROM " & Chr(34) & vbNewLine
        ' here are the Tables
        For I = LBound(strTables) To UBound(strTables)
            ' clean out all CRs and LFs and such
            strTemp = Replace(strTables(I), Chr(13), "")
            strTemp = Replace(strTemp, Chr(10), "")
            strTemp = Replace(strTemp, vbCrLf, "")
            strTemp = Replace(strTemp, vbNewLine, "")
            strTemp = Replace(strTemp, vbTab, "")
            strTemp = Trim$(strTemp)
                If I = UBound(strTables) Then
                    If Trim$(strTables(I)) = "" Then
                        ' skip it
                    Else
                        strOutPut = strOutPut & "strSQL = strSQL & " & Chr(34) & strTemp & ", " & Chr(34) & vbNewLine
                    End If
                Else
                    strOutPut = strOutPut & "strSQL = strSQL & " & Chr(34) & strTemp & ", " & Chr(34) & vbNewLine
                End If
        Next
        ' here are the Where clauses
        For I = LBound(strValuePart) To UBound(strValuePart)
            strTemp = Trim$(strValuePart(I))
            ' get rid of embedded double quotes
            strTemp = Replace(strTemp, Chr(34), Chr(39))
            ' get rid of parentheses
            strTemp = Replace(strTemp, ")", "")
            strTemp = Replace(strTemp, "(", "")
            ' fix up our equality symbols
            strTemp = Replace(strTemp, "=", " = ")
            strTemp = Replace(strTemp, "> =", " >= ")
            strTemp = Replace(strTemp, "< =", " <= ")
            ' clean out double spaces
            strTemp = Replace(strTemp, "  ", " ")
                If Trim$(strTemp) = "" Then
                    ' skip empties
                Else
                    strOutPut = strOutPut & "strSQL = strSQL & " & Chr(34) & strTemp & " " & Chr(34) & vbNewLine
                End If
        Next
    strOutPut = Replace(strOutPut, ";", "")
    ' fix unnecessary comma before WHERE, if it exists (artifact from how we get table names)
    If InStr(1, strOutPut, ", " & Chr(34) & vbNewLine & "strSQL = strSQL & " & Chr(34) & "WHERE ", vbTextCompare) Then
        strOutPut = Replace(strOutPut, ", " & Chr(34) & vbNewLine & "strSQL = strSQL & " & Chr(34) & "WHERE ", " " & Chr(34) & vbNewLine & "strSQL = strSQL & " & Chr(34) & "WHERE ", , , vbTextCompare)
    End If
    GetSelect = strOutPut

Exit_GetSelect:
    
    On Error Resume Next
    Erase strTables
    Erase strFieldList
    Erase strValuePart
    On Error GoTo 0
    Exit Function

Err_GetSelect:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In modBuildSQL, during GetSelect" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_GetSelect
    End Select
    
End Function
Public Function GetUpdate(strTheSQLStatement As String) As String
Attribute GetUpdate.VB_UserMemId = 1610612741
   
    ' Purpose   :
    ' Parameters: strTheSQLStatement
    ' Returns   : String
    ' Modified  : 3/27/2002 By BB

    Dim strAllButFirst As String
    Dim strBeginning   As String
    Dim strTableName   As String
    Dim strFields      As String
    Dim strValues      As String
    Dim strFieldPart() As String
    Dim strFieldList() As String
    Dim strValuePart() As String
    Dim strValueList() As String
    Dim strOutPut      As String
    Dim strTemp        As String
    Dim iPos           As Integer
    Dim iSpacePos      As Integer
    Dim I              As Integer

    On Error GoTo Err_GetUpdate
    
    
    ' the incoming string will be something like...
    ' update TABLENAME SET (FIELD1 = VALUE1, FIELD2 = VALUE2, etc)
    
    ' UPDATE is 6 characters long, so...
    strAllButFirst = Trim$(Right$(strTheSQLStatement, Len(strTheSQLStatement) - 6))
    ' get rid of any WHERE clause...
    iPos = InStr(1, strAllButFirst, " where ", vbTextCompare)
        If iPos > 0 Then
            strAllButFirst = Trim$(Left$(strAllButFirst, iPos - 1))
        End If
    ' extract TABLE NAME
    iSpacePos = InStr(strAllButFirst, " ")
    strTableName = Trim$(Left$(strAllButFirst, iSpacePos))
    strAllButFirst = Trim$(Right$(strAllButFirst, Len(strAllButFirst) - iSpacePos))
    ' extract TABLE NAME
    iSpacePos = InStr(strAllButFirst, " ")
    strAllButFirst = Replace(strAllButFirst, strTableName, "")
    strAllButFirst = Trim$(Replace(strAllButFirst, " set ", " SET ", vbTextCompare))
    iPos = InStr(1, strAllButFirst, "SET ", vbTextCompare)
    strFields = Trim$(Right$(strAllButFirst, Len(strAllButFirst) - 4))
    strFields = Replace(strFields, "(", "")
    strFields = Replace(strFields, ")", "")
    strFieldPart = Split(strFields, ",")
    ReDim strFieldList(UBound(strFieldPart))
        For I = LBound(strFieldPart) To UBound(strFieldPart)
            iPos = InStr(strFieldPart(I), "=")
                If iPos > 0 Then
                    strTemp = Trim$(Left$(strFieldPart(I), iPos - 1))
                    strFieldList(I) = strTemp
                Else
                    strFieldList = Split(strFieldPart(I), "=")
                End If
        Next
    strValues = Trim$(Right$(strAllButFirst, Len(strAllButFirst) - 4))
    strValues = Replace(strValues, "(", "")
    strValues = Replace(strValues, ")", "")
    strValuePart = Split(strValues, ",")
    ReDim strValueList(UBound(strValuePart))
        For I = LBound(strValuePart) To UBound(strValuePart)
            iPos = InStr(strFieldPart(I), "=")
                If iPos > 0 Then
                    strTemp = Trim$(Right$(strValuePart(I), iPos - 1))
                    strValueList(I) = strTemp
                Else
                    strValueList = Split(strValuePart(I), "=")
                End If
        Next
    strOutPut = ""
    strOutPut = "strSQL = " & Chr(34) & "UPDATE " & Chr(34) & vbNewLine
    strOutPut = strOutPut & "strSQL = strSQL & " & Chr(34) & strTableName & " " & Chr(34) & vbNewLine
    strOutPut = strOutPut & "strSQL = strSQL & " & Chr(34) & "SET " & Chr(34) & vbNewLine
    strOutPut = strOutPut & "strSQL = strSQL & " & Chr(34) & "(" & Chr(34) & vbNewLine
        For I = LBound(strFieldList) To UBound(strFieldList)
            If I = UBound(strFieldList) Then
                strOutPut = strOutPut & "strSQL = strSQL & " & Chr(34) & Trim$(strFieldList(I)) & " = " & Trim$(strValueList(I)) & Chr(34) & vbNewLine
            Else
                strOutPut = strOutPut & "strSQL = strSQL & " & Chr(34) & Trim$(strFieldList(I)) & " = " & Trim$(strValueList(I)) & ", " & Chr(34) & vbNewLine
            End If
        Next
    strOutPut = strOutPut & "strSQL = strSQL & " & Chr(34) & ") " & Chr(34) & vbNewLine
    GetUpdate = strOutPut

Exit_GetUpdate:
    
    On Error Resume Next
    Erase strFieldList
    Erase strValueList
    On Error GoTo 0
    Exit Function

Err_GetUpdate:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In modBuildSQL, during GetUpdate" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_GetUpdate
    End Select

End Function
Public Function HandleWhere(strSQLWithWhere As String) As String
Attribute HandleWhere.VB_UserMemId = 1610612742
   
    ' Purpose   : parse and format the Where clause
    ' Parameters: strSQLWithWhere
    ' Returns   : String
    ' Modified  : 3/27/2002 By BB
    '
    ' WHERE Field1 = Value1 AND Field2 = Value2 etc
    ' WHERE Field1 = Value1 OR  Field2 = Value2 etc
    ' so many possibilities...let's just look for the words WHERE, AND and OR
    ' and split the string accordingly
    
    Dim iPos              As Integer
    Dim I                 As Integer
    Dim strAllButWhere    As String
    Dim strOutPut         As String
    Dim strCriteriaList() As String
    Dim strValueList()    As String
    Dim strTmp            As String
    
    On Error GoTo Err_HandleWhere
    
    'SELECT GroupName,UserFirstName FROM tblGRPUserGroups  WHERE GroupName = Tech Group Work Orders AND UserFirstName >= Art
    
    iPos = InStr(1, strSQLWithWhere, " where ", vbTextCompare)
        If iPos > 0 Then
            strAllButWhere = Trim$(Right$(strSQLWithWhere, Len(strSQLWithWhere) - (iPos + 5)))
        Else
            HandleWhere = ""
            GoTo Exit_HandleWhere
        End If
    ' reset iPos variable
    iPos = 0
    For I = 1 To Len(strAllButWhere) - 1
        If Mid$(strAllButWhere, I, 1) = "=" Then
            iPos = iPos + 1
        End If
    Next
    strOutPut = ""
    strOutPut = "strSQL = strSQL & " & Chr(34) & "WHERE " & Chr(34) & vbNewLine
    If InStr(1, strAllButWhere, " and ", vbTextCompare) Then
        strTmp = GetStringBetween(strAllButWhere, "", " and ")
        'Debug.Print "strTmp: " & strTmp
        If InStr(1, strTmp, "=", vbTextCompare) Then
            If InStr(1, strTmp, ".", vbTextCompare) Then
                ' must be an Access field/table name, don't need quotes if it's in square brackets
            ElseIf InStr(1, strTmp, "].[", vbTextCompare) Then
                ' must be an Access field/table name, don't need quotes if it's in square brackets
            Else
                If InStr(1, strTmp, "= ", vbTextCompare) Then
                    strTmp = Replace(strTmp, "= ", "= '", , , vbTextCompare)
                    strTmp = strTmp & "' "
                ElseIf InStr(1, strTmp, "=", vbTextCompare) Then
                    strTmp = Replace(strTmp, "=", "= '", , , vbTextCompare)
                    strTmp = strTmp & "' "
                End If
            End If
        End If
        Debug.Print "strTmp: " & strTmp
    End If
    If InStr(1, strAllButWhere, " or ", vbTextCompare) Then
        strTmp = GetStringBetween(strAllButWhere, "", " or ")
        Debug.Print "strTmp: " & strTmp
    End If
    strOutPut = strOutPut & "strSQL = strSQL & " & Chr(34) & " AND " & Chr(34) & vbNewLine
    strOutPut = strOutPut & "strSQL = strSQL & " & Chr(34) & strTmp & " " & Chr(34) & vbNewLine
    Debug.Print "strOutPut: " & strOutPut
    ' 2 = signs in clause
    If iPos = 1 Then
        strOutPut = strOutPut & "strSQL = strSQL & " & Chr(34) & strAllButWhere & " " & Chr(34) & vbNewLine
        HandleWhere = strOutPut
        GoTo Exit_HandleWhere
    End If
    ' stuff an AND in there
    strOutPut = strOutPut & "strSQL = strSQL & " & Chr(34) & "AND " & Chr(34) & vbNewLine
    ' stuff an OR in there
    strOutPut = strOutPut & "strSQL = strSQL & " & Chr(34) & "OR " & Chr(34) & vbNewLine
    strOutPut = strOutPut & "strSQL = strSQL & " & Chr(34) & "(" & Chr(34) & vbNewLine
'            For I = LBound(strCriteriaList) To UBound(strCriteriaList)
'                If I = UBound(strCriteriaList) Then
'                    strOutPut = strOutPut & "strSQL = strSQL & " & Chr(34) & Trim$(strCriteriaList(I)) & " = " & Trim$(strValueList(I)) & Chr(34) & vbNewLine
'                Else
'                    strOutPut = strOutPut & "strSQL = strSQL & " & Chr(34) & Trim$(strCriteriaList(I)) & " = " & Trim$(strValueList(I)) & ", " & Chr(34) & vbNewLine
'                End If
'            Next
    strOutPut = strOutPut & "strSQL = strSQL & " & Chr(34) & ")" & Chr(34) & vbNewLine
    
    HandleWhere = strOutPut

Exit_HandleWhere:
    
    On Error Resume Next
    Erase strCriteriaList
    Erase strValueList
    On Error GoTo 0
    Exit Function

Err_HandleWhere:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In modBuildSQL, during HandleWhere" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_HandleWhere
    End Select

End Function
Public Function SearchString(strSearchText As String) As String
Attribute SearchString.VB_UserMemId = 1610612737
    Dim astrKeyWords() As String
    Dim strQuery       As String
    Dim C              As Integer
    
    strSearchText = Replace(strSearchText, Chr(34) & Chr(34), Chr(34), 1)
    astrKeyWords = Split(strSearchText, Chr(32), -1, 1)
        For C = 0 To UBound(astrKeyWords)
            Select Case LCase(astrKeyWords(C))
                Case "and", "+"
                    If LCase(astrKeyWords(C + 1)) <> "not" Then strQuery = strQuery & " AND "
                Case "or"
                    strQuery = strQuery & " OR "
                Case "not", "-"
                    strQuery = strQuery & " AND NOT "
                Case "near"
                    strQuery = strQuery & " NEAR "
                Case Else
                    If (astrKeyWords(C) <> "") Then
                        If (C > 0) And (astrKeyWords(C) <> "") Then
                                If (Left(astrKeyWords(C - 1), 1) = Chr(34)) And (Right(astrKeyWords(C - 1), 1) <> Chr(34)) And (Right(astrKeyWords(C), 1) <> Chr(34)) Then
                                    astrKeyWords(C) = astrKeyWords(C) & Chr(34)
                                End If
                                If (Left(astrKeyWords(C - 1), 1) = Chr(34)) And (Right(astrKeyWords(C - 1), 1) = Chr(34)) Then
                                    astrKeyWords(C) = " AND " & astrKeyWords(C)
                                End If
                                If (Left(astrKeyWords(C - 1), 1) <> Chr(34)) And (Right(astrKeyWords(C), 1) = Chr(34)) Then
                                    astrKeyWords(C) = Replace(astrKeyWords(C), Chr(34), "", 1)
                                End If
                                If (LCase(astrKeyWords(C - 1)) <> "and") And (LCase(astrKeyWords(C - 1)) <> "or") And (LCase(astrKeyWords(C - 1)) <> "not") And (LCase(astrKeyWords(C - 1)) <> "near") And (LCase(astrKeyWords(C - 1)) <> "") And (Left(astrKeyWords(C - 1), 1) <> Chr(34)) Then
                                    strQuery = strQuery & " AND " & astrKeyWords(C)
                                Else
                                    strQuery = strQuery & astrKeyWords(C) & Chr(32)
                                End If
                        End If
                        If (C = 0) Then strQuery = astrKeyWords(C) & Chr(32)
                    End If
            End Select
        Next
        Do While (Right(strQuery, 1) = Chr(32)) And (Len(strQuery) > 0)
            strQuery = Left(strQuery, Len(strQuery) - 1)
        Loop
    SearchString = strQuery
    Erase astrKeyWords
End Function
Public Function SplitTheWhere(strWhereClause As String) As String
Attribute SplitTheWhere.VB_UserMemId = 1610612743
   
    Dim iPos   As Integer
    Dim strTmp As String
    
    On Error GoTo Err_SplitTheWhere
    
    strTmp = Replace(strWhereClause, "where ", "WHERE,", , , vbTextCompare)
    strTmp = Replace(strTmp, " and ", ",AND, ", , , vbTextCompare)
    strTmp = Replace(strTmp, " or ", ",OR, ", , , vbTextCompare)
    SplitTheWhere = strTmp
    GoTo Exit_SplitTheWhere
    
    iPos = InStr(1, strWhereClause, "where ", vbTextCompare)
        Debug.Print "iPos: " & iPos
        If iPos > 0 Then
            If iPos = 1 Then
                strTmp = Trim$(Mid$(strWhereClause, 6, Len(strWhereClause)))
            End If
            'strTmp = Right$(strWhereClause, Len(strWhereClause) - iPos)
            Debug.Print "strTmp: " & strTmp
        End If
    iPos = InStr(1, strWhereClause, " and ", vbTextCompare)
        If iPos > 0 Then
            strTmp = strTmp & "," & Right$(strWhereClause, Len(strWhereClause) - iPos)
            Debug.Print "strTmp: " & strTmp
        End If
    iPos = InStr(1, strWhereClause, " or ", vbTextCompare)
        If iPos > 0 Then
            strTmp = strTmp & "," & Right$(strWhereClause, Len(strWhereClause) - iPos)
            Debug.Print "strTmp: " & strTmp
        End If
    SplitTheWhere = "WHERE," & strTmp

Exit_SplitTheWhere:
    
    On Error GoTo 0
    Exit Function

Err_SplitTheWhere:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In modParseSQL, during SplitTheWhere" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_SplitTheWhere
    End Select

End Function
