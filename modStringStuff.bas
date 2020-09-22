Attribute VB_Name = "modStringStuff"
' StringModule for Visual Basic
Option Explicit

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Function Lines(MyStr As String) As Integer
    Dim CurNewLine As Integer
    
    CurNewLine = 1
    If (Len(MyStr)) Then
        Lines = 1
    Else
        Lines = 0
        Exit Function
    End If
    Do
        CurNewLine = InStr(CurNewLine + 1, MyStr, vbNewLine, vbBinaryCompare)
        If (CurNewLine) Then Lines = Lines + 1
    Loop Until CurNewLine = 0
End Function
Function Line(MyStr As String, WhichLine As Integer) As String
    Dim CurNewLine As Integer
    Dim CurLine    As Integer
    Dim CurEndLine As Integer
    
    CurNewLine = 1
        If (Len(MyStr)) Then
            CurLine = 1
        Else
            Line = ""
            Exit Function
        End If
        While (CurLine < WhichLine And CurNewLine <> 0)
            CurNewLine = Len(vbNewLine) + InStr(CurNewLine, MyStr, vbNewLine, vbBinaryCompare)
            If (CurNewLine) Then CurLine = CurLine + 1
        Wend
    CurEndLine = InStr(CurNewLine, MyStr, vbNewLine, vbBinaryCompare)
        If CurEndLine = 0 Then
            Line = Mid(MyStr, CurNewLine)
        Else
            Line = Mid(MyStr, CurNewLine, CurEndLine - CurNewLine)
        End If
End Function
Function RemoveLine(MyStr As String, WhichLine As Integer) As String
    Dim CurNewLine As Integer
    Dim CurLine    As Integer
    Dim CurEndLine As Integer
    
        If WhichLine < 1 Or WhichLine > Lines(MyStr) Then
            RemoveLine = MyStr
            Exit Function
        End If
    CurNewLine = 1
        If (Len(MyStr)) Then
            CurLine = 1
        Else
            RemoveLine = ""
            Exit Function
        End If
        While (CurLine < WhichLine And CurNewLine <> 0)
            CurNewLine = Len(vbNewLine) + InStr(CurNewLine, MyStr, vbNewLine, vbBinaryCompare)
            If (CurNewLine) Then CurLine = CurLine + 1
        Wend
    CurEndLine = InStr(CurNewLine, MyStr, vbNewLine, vbBinaryCompare)
        If CurEndLine = 0 Then
            MyStr = Left(MyStr, Max(CurNewLine - Len(vbNewLine) - 1, 0))
        Else
            MyStr = Left(MyStr, Max(CurNewLine - Len(vbNewLine) - 1, 0)) & Mid(MyStr, CurEndLine)
            If WhichLine = 1 And Left(MyStr, Len(vbNewLine)) = vbNewLine Then MyStr = Mid(MyStr, Len(vbNewLine) + 1)
        End If
    RemoveLine = MyStr
End Function
Function AddLine(MyStr As String, LineToAdd As String, Num As Integer) As String
    Dim CurNewLine As Integer
    Dim CurLine    As Integer
    
        If Num > Lines(MyStr) + 1 Then
            AddLine = MyStr
            Exit Function
        End If
    CurNewLine = 1
        If (Len(MyStr)) Then
            CurLine = 1
        Else
            MyStr = LineToAdd
            AddLine = MyStr
            Exit Function
        End If
        If Num = 1 Then
            'Insert first
            MyStr = LineToAdd & vbNewLine & MyStr
        Else
            If (Num = 0) Or (Num = Lines(MyStr) + 1) Then
                'Insert last
                MyStr = MyStr & vbNewLine & LineToAdd
            Else
                'Insert in the middle
                While (CurLine < Num And CurNewLine <> 0)
                    CurNewLine = Len(vbNewLine) + InStr(CurNewLine, MyStr, vbNewLine, vbBinaryCompare)
                    If (CurNewLine) Then CurLine = CurLine + 1
                Wend
                MyStr = Left(MyStr, CurNewLine - 1) & LineToAdd & vbNewLine & Mid(MyStr, CurNewLine)
            End If
        End If
    AddLine = MyStr
End Function
Function Word(Line As String, Num As Integer) As String
    Dim I           As Integer
    Dim CurWord     As Integer
    Dim Beg_CurWord As Integer
    Dim End_CurWord As Integer
    
    I = 1
        If (Num = 0) Then Word = "": Exit Function
        Do
                While (((Mid(Line, I, 1) = " ") Or (Mid(Line, I, 1) = vbTab) Or (Mid(Line, I, 2) = vbNewLine)) And I <= Len(Line))
                    If Mid(Line, I, 2) = vbNewLine Then
                        I = I + 2
                    Else
                        I = I + 1
                    End If
                Wend
            Beg_CurWord = I
            CurWord = CurWord + 1
                While (((Mid(Line, I, 1) <> " ") And (Mid(Line, I, 1) <> vbTab) And (Mid(Line, I, 2) <> vbNewLine)) And I <= Len(Line))
                    I = I + 1
                Wend
            End_CurWord = I
        Loop Until (CurWord = Num Or I = Len(Line))
    Word = Mid(Line, Beg_CurWord, End_CurWord - Beg_CurWord)
End Function
Function ToWord(Line As String, Num As Integer) As String
    Dim I           As Integer
    Dim CurWord     As Integer
    Dim Beg_CurWord As Integer
    Dim End_CurWord As Integer
    
    I = 1
        If (Num = 0) Then
            ToWord = ""
            Exit Function
        End If
        Do
                While (((Mid(Line, I, 1) = " ") Or (Mid(Line, I, 1) = vbTab) Or (Mid(Line, I, 2) = vbNewLine)) And I <= Len(Line))
                    If Mid(Line, I, 2) = vbNewLine Then
                        I = I + 2
                    Else
                        I = I + 1
                    End If
                Wend
            Beg_CurWord = I
            CurWord = CurWord + 1
                While (((Mid(Line, I, 1) <> " ") And (Mid(Line, I, 1) <> vbTab) And (Mid(Line, I, 2) <> vbNewLine)) And I <= Len(Line))
                    I = I + 1
                Wend
            End_CurWord = I
        Loop Until (CurWord = Num Or I = Len(Line))
    ToWord = Left(Line, End_CurWord)
End Function
Function FromWord(Line As String, Num As Integer) As String
    Dim I           As Integer
    Dim CurWord     As Integer
    Dim Beg_CurWord As Integer
    Dim End_CurWord As Integer
    
    I = 1
        If (Num = 0) Then FromWord = Line: Exit Function
        Do
                While (((Mid(Line, I, 1) = " ") Or (Mid(Line, I, 1) = vbTab) Or (Mid(Line, I, 2) = vbNewLine)) And I <= Len(Line))
                    If Mid(Line, I, 2) = vbNewLine Then
                        I = I + 2
                    Else
                        I = I + 1
                    End If
                Wend
            Beg_CurWord = I
            CurWord = CurWord + 1
                While (((Mid(Line, I, 1) <> " ") And (Mid(Line, I, 1) <> vbTab) And (Mid(Line, I, 2) <> vbNewLine)) And I <= Len(Line))
                    I = I + 1
                Wend
            End_CurWord = I
        Loop Until (CurWord = Num Or I = Len(Line))
    FromWord = Mid(Line, Beg_CurWord)
End Function
Function Words(Line As String) As Integer
    Dim I       As Integer
    Dim CurWord As Integer
    
    I = 1
        Do
            While (((Mid(Line, I, 1) = " ") Or (Mid(Line, I, 1) = vbTab) Or (Mid(Line, I, 2) = vbNewLine)) And I <= Len(Line))
                    If Mid(Line, I, 2) = vbNewLine Then
                        I = I + 2
                    Else
                        I = I + 1
                    End If
            Wend
            If (I <= Len(Line)) Then CurWord = CurWord + 1
            While (((Mid(Line, I, 1) <> " ") And (Mid(Line, I, 1) <> vbTab) And (Mid(Line, I, 2) <> vbNewLine)) And I <= Len(Line))
                I = I + 1
            Wend
        Loop While (I <= Len(Line))
    Words = CurWord
End Function
Function ReplaceIt(myString As String, RepMe As String, WithMe As String, Optional StartAt As Long = 1)
    Dim I As Integer
    
    I = InStr(StartAt, myString, RepMe, vbTextCompare)
    If I > 0 Then
        I = I - 1
        ReplaceIt = Left(myString, I) & WithMe & Right(myString, Len(myString) - Len(RepMe) - I)
    End If
End Function
Function ReplaceAll(myString As String, RepMe As String, WithMe As String, Optional StartAt As Long = 1) As String
    ReplaceAll = myString
    StartAt = InStr(StartAt, ReplaceAll, RepMe, vbTextCompare)
    While StartAt
        ReplaceAll = Replace(ReplaceAll, RepMe, WithMe, StartAt)
        StartAt = StartAt + Len(WithMe) + 1
        StartAt = InStr(StartAt, ReplaceAll, RepMe, vbTextCompare)
    Wend
End Function
Function strIsPrefix(MyStr As String, MyPrefix As String) As Boolean
    If Left(MyStr, Len(MyPrefix)) = MyPrefix Then
        strIsPrefix = True
    Else
        strIsPrefix = False
    End If
End Function
Function strIsSuffix(MyStr As String, MySuffix As String) As Boolean
    If Right(MyStr, Len(MySuffix)) = MySuffix Then
        strIsSuffix = True
    Else
        strIsSuffix = False
    End If
End Function
Function Max(A As Long, B As Long) As Long
   
    On Error GoTo Err_Max
    
    If A > B Then
        Max = A
    Else
        Max = B
    End If

Exit_Max:
    
    On Error GoTo 0
    Exit Function

Err_Max:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In modStringStuff, during Max" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_Max
    End Select
    
End Function
Public Sub ShowHelp()
   
    On Error GoTo Err_ShowHelp
    
    On Error Resume Next
    
    frmHelp.Show vbModal

Exit_ShowHelp:
    
    On Error GoTo 0
    Exit Sub

Err_ShowHelp:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In modStringStuff, during ShowHelp" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_ShowHelp
    End Select
    
End Sub
Public Function TrimNulls(strIn As String) As String
   
    ' Comments  : Removes the null terminator from a string
    ' Parameters: strIn - String to modify
    ' Returns   : Modified string
    
    Dim lngChr As Long
    
    On Error GoTo Err_TrimNulls
    
    lngChr = InStr(strIn, Chr$(0))
    If lngChr > 0 Then
        TrimNulls = Left$(strIn, lngChr - 1)
    Else
        TrimNulls = strIn
    End If

Exit_TrimNulls:
    
    On Error GoTo 0
    Exit Function

Err_TrimNulls:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In modStringStuff, during TrimNulls" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_TrimNulls
    End Select
    
End Function
