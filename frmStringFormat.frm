VERSION 5.00
Begin VB.Form frmStringFormat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "    VBRSGen - BB's VB Recordset Generator"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   570
   ClientWidth     =   10095
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmStringFormat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   10095
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBack 
      Caption         =   "Select Again"
      Height          =   285
      Left            =   7935
      TabIndex        =   12
      Top             =   615
      Width           =   1305
   End
   Begin VB.CommandButton CmdClearAll 
      Caption         =   "Clear All"
      Height          =   285
      Left            =   7935
      TabIndex        =   8
      Top             =   315
      Width           =   1305
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   270
      Left            =   9300
      TabIndex        =   7
      Top             =   300
      Width           =   765
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   270
      Left            =   9300
      TabIndex        =   6
      Top             =   15
      Width           =   765
   End
   Begin VB.CommandButton cmdFormat 
      Caption         =   "Generate Code"
      Height          =   285
      Left            =   7935
      TabIndex        =   5
      Top             =   15
      Width           =   1305
   End
   Begin VB.Frame fOptions 
      Caption         =   "  Options  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   2895
      TabIndex        =   4
      Top             =   15
      Width           =   4020
      Begin VB.Frame fOutput 
         Caption         =   "  Output  "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   1845
         TabIndex        =   13
         Top             =   195
         Width           =   2115
         Begin VB.OptionButton optClipBoard 
            Caption         =   "Clipboard"
            Height          =   210
            Left            =   105
            TabIndex        =   15
            Top             =   225
            Value           =   -1  'True
            Width           =   1065
         End
         Begin VB.OptionButton optFile 
            Caption         =   "Notepad"
            Height          =   210
            Left            =   1155
            TabIndex        =   14
            Top             =   225
            Width           =   915
         End
      End
      Begin VB.Frame fraDaoAdo 
         Caption         =   "  Data Access Type  "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   60
         TabIndex        =   9
         Top             =   210
         Width           =   1770
         Begin VB.OptionButton optADO 
            Caption         =   "ADO"
            Height          =   210
            Left            =   105
            TabIndex        =   11
            Top             =   225
            Value           =   -1  'True
            Width           =   645
         End
         Begin VB.OptionButton optDAO 
            Caption         =   "DAO"
            Height          =   210
            Left            =   810
            TabIndex        =   10
            Top             =   240
            Width           =   645
         End
      End
   End
   Begin VB.TextBox txtNewString 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4890
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   2955
      Width           =   9990
   End
   Begin VB.TextBox txtOldString 
      Height          =   1650
      Left            =   45
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1065
      Width           =   9990
   End
   Begin VB.Label lblNewString 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Generated VB Code"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   2730
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label lOldText 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "String to be Transformed into VB ADO or DAO Code"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   2
      Top             =   825
      Width           =   3450
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileOk 
         Caption         =   "OK"
      End
      Begin VB.Menu mnuFileSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileCancel 
         Caption         =   "Cancel"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Data Access Options"
      Begin VB.Menu mnuOptionsDataAccessTypeADO 
         Caption         =   "ADO"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuOptionsDataAccessTypeDAO 
         Caption         =   "DAO"
      End
   End
   Begin VB.Menu mnuAction 
      Caption         =   "Action"
      Begin VB.Menu mnuGenerateVBCode 
         Caption         =   "Generate VB Code"
      End
      Begin VB.Menu mnuClearAll 
         Caption         =   "Clear All"
      End
      Begin VB.Menu mnuSelectAgain 
         Caption         =   "Select Again"
      End
   End
   Begin VB.Menu mnuOutput 
      Caption         =   "Output"
      Begin VB.Menu mnuOutputClipboard 
         Caption         =   "Clipboard"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuOutputNotepad 
         Caption         =   "Notepad"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpHelp 
         Caption         =   "Help"
      End
      Begin VB.Menu mnuHelpSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About..."
      End
   End
End
Attribute VB_Name = "frmStringFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DIM_STR1    As String = "Dim "
Private Const DIM_STR2    As String = " As String"
Private Const CONT_STR    As String = " & _"
Private Const CONNECT_STR As String = " & "

Private Const SELECT_STR  As String = "SELECT "
Private Const FROM_STR    As String = " FROM "
Private Const WHERE_STR   As String = " WHERE "
Private Const GROUPBY_STR As String = " GROUP BY "
Private Const UPDATE_STR  As String = "UPDATE "
Private Const INSERT_STR  As String = "INSERT INTO "
Private Const DELETE_STR  As String = "DELETE "

Private aSQLVar()         As Integer
Private Clicked           As Boolean
Private mbUseADO          As Boolean
Private mbClipboard       As Boolean
Private mbFile            As Boolean

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Function AppendString(sVar As String, sLine As String, bEnd As Boolean) As String
   
    On Error GoTo Err_AppendString
    
    If bEnd Then
        AppendString = sVar & " = " & sVar & CONNECT_STR & _
            Chr(34) & sLine & Chr(34)
    Else
        AppendString = sVar & " = " & CONNECT_STR & _
            Chr(34) & sLine & Chr(34)
        'AppendString = sVar & " = " & sVar & CONNECT_STR & _
            Chr(34) & sLine & Chr(34)
    End If

Exit_AppendString:
    
    On Error GoTo 0
    Exit Function

Err_AppendString:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmStringFormat, during AppendString" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_AppendString
    End Select

End Function
Function CleanString(szOriginal)
Attribute CleanString.VB_UserMemId = 1610809345
 
    On Error GoTo Err_CleanString
    
    If szOriginal = "" Then
        CleanString = "NULL"
    Else
        CleanString = Substitute(szOriginal, "'", "''")
        CleanString = Substitute(CleanString, "’", "’’")
    End If

Exit_CleanString:
    
    On Error GoTo 0
    Exit Function

Err_CleanString:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmStringFormat, during CleanString" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_CleanString
    End Select
    
End Function
Public Sub ClipboardCopy(Text As String)
Attribute ClipboardCopy.VB_UserMemId = 1610809354
   
    'Copies text to the clipboard
    
    On Error GoTo Err_ClipboardCopy
    
    Clipboard.Clear
    Clipboard.SetText Text

Exit_ClipboardCopy:
    
    On Error GoTo 0
    Exit Sub

Err_ClipboardCopy:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmStringFormat, during ClipboardCopy" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_ClipboardCopy
    End Select
    
End Sub
Private Sub cmdBack_Click()
       
    On Error GoTo Err_cmdBack_Click
    
    SelectAgain

Exit_cmdBack_Click:
    
    On Error GoTo 0
    Exit Sub

Err_cmdBack_Click:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmStringFormat, during cmdBack_Click" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_cmdBack_Click
    End Select
    
End Sub
Private Sub cmdCancel_Click()
   
    On Error GoTo Err_cmdCancel_Click
    
    Dim F As Form
    
    On Error Resume Next
    
    For Each F In Forms
        Unload F
    Next
    Unload Me

Exit_cmdCancel_Click:
    
    On Error GoTo 0
    Exit Sub

Err_cmdCancel_Click:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmStringFormat, during cmdCancel_Click" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_cmdCancel_Click
    End Select
    
End Sub
Private Sub CmdClearAll_Click()
   
    On Error GoTo Err_CmdClearAll_Click
    
    ClearAll
    
Exit_CmdClearAll_Click:
    
    On Error GoTo 0
    Exit Sub

Err_CmdClearAll_Click:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmStringFormat, during CmdClearAll_Click" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_CmdClearAll_Click
    End Select

End Sub
Private Sub cmdFormat_Click()
   
    On Error GoTo Err_cmdFormat_Click
    
    FormatString
        
Exit_cmdFormat_Click:
    
    On Error GoTo 0
    Exit Sub

Err_cmdFormat_Click:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmStringFormat, during cmdFormat_Click" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_cmdFormat_Click
    End Select
    
End Sub
Private Sub cmdOK_Click()
    
    On Error GoTo Err_cmdOK_Click
        
    Finish
    
Exit_cmdOK_Click:
    
    Unload Me
    On Error GoTo 0
    Exit Sub

Err_cmdOK_Click:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmStringFormat, during cmdOK_Click" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_cmdOK_Click
    End Select

End Sub
Private Function ContinueString(sLine As String, bEnd As Boolean) As String
   
    On Error GoTo Err_ContinueString
    
    If bEnd Then
        ContinueString = Chr(34) & sLine & Chr(34)
    Else
        ContinueString = Chr(34) & sLine & Chr(34) & CONT_STR
    End If

Exit_ContinueString:
    
    On Error GoTo 0
    Exit Function

Err_ContinueString:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmStringFormat, during ContinueString" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_ContinueString
    End Select
    
End Function
Private Sub Form_Load()
   
    ' use ADO as default
    
    On Error GoTo Err_Form_Load
    
    mbUseADO = True

Exit_Form_Load:
    
    On Error GoTo 0
    Exit Sub

Err_Form_Load:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmStringFormat, during Form_Load" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_Form_Load
    End Select
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
   
    Dim F As Form
    
    On Error GoTo Err_Form_Unload
    
    On Error Resume Next
    
    For Each F In Forms
        Unload F
    Next

Exit_Form_Unload:
    
    On Error Resume Next
        If Not (F Is Nothing) Then
            Set F = Nothing
        End If
    On Error GoTo 0
    Exit Sub

Err_Form_Unload:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmStringFormat, during Form_Unload" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_Form_Unload
    End Select

End Sub
Private Sub mnuClearAll_Click()
   
    On Error GoTo Err_mnuClearAll_Click
    
    ClearAll

Exit_mnuClearAll_Click:
    
    On Error GoTo 0
    Exit Sub

Err_mnuClearAll_Click:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmStringFormat, during mnuClearAll_Click" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_mnuClearAll_Click
    End Select
    
End Sub
Private Sub mnuFileCancel_Click()
   
    Dim F As Form
    
    On Error GoTo Err_mnuFileCancel_Click
    
    On Error Resume Next
    
    For Each F In Forms
        Unload F
    Next
    
Exit_mnuFileCancel_Click:
    
    On Error Resume Next
        If Not (F Is Nothing) Then
            Set F = Nothing
        End If
    On Error GoTo 0
    Unload Me
    Exit Sub

Err_mnuFileCancel_Click:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmStringFormat, during mnuFileCancel_Click" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_mnuFileCancel_Click
    End Select
    
End Sub
Private Sub mnuFileOk_Click()
   
    On Error GoTo Err_mnuFileOk_Click
    
    Finish
    
Exit_mnuFileOk_Click:
    
    On Error Resume Next
    Unload Me
    On Error GoTo 0
    Exit Sub

Err_mnuFileOk_Click:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmStringFormat, during mnuFileOk_Click" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_mnuFileOk_Click
    End Select
    
End Sub
Private Sub mnuGenerateVBCode_Click()
   
    On Error GoTo Err_mnuGenerateVBCode_Click
    
    FormatString

Exit_mnuGenerateVBCode_Click:
    
    On Error GoTo 0
    Exit Sub

Err_mnuGenerateVBCode_Click:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmStringFormat, during mnuGenerateVBCode_Click" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_mnuGenerateVBCode_Click
    End Select
    
End Sub
Private Sub mnuHelpAbout_Click()
   
    On Error GoTo Err_mnuHelpAbout_Click
    
    ShowHelp

Exit_mnuHelpAbout_Click:
    
    On Error GoTo 0
    Exit Sub

Err_mnuHelpAbout_Click:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmStringFormat, during mnuHelpAbout_Click" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_mnuHelpAbout_Click
    End Select
    
End Sub
Private Sub mnuHelpHelp_Click()
   
    On Error GoTo Err_mnuHelpHelp_Click
    
    frmHelp.Show

Exit_mnuHelpHelp_Click:
    
    On Error GoTo 0
    Exit Sub

Err_mnuHelpHelp_Click:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmStringFormat, during mnuHelpHelp_Click" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_mnuHelpHelp_Click
    End Select
    
End Sub
Private Sub mnuOptionsDataAccessTypeADO_Click()
   
    On Error GoTo Err_mnuOptionsDataAccessTypeADO_Click
    
    mnuOptionsDataAccessTypeADO.Checked = Not mnuOptionsDataAccessTypeADO.Checked
    mbUseADO = mnuOptionsDataAccessTypeADO.Checked
    mnuOptionsDataAccessTypeDAO.Checked = Not mbUseADO
    optADO.Value = mnuOptionsDataAccessTypeADO.Checked
    optDAO.Value = Not mnuOptionsDataAccessTypeADO.Checked
    FormatString
    
Exit_mnuOptionsDataAccessTypeADO_Click:
    
    On Error GoTo 0
    Exit Sub

Err_mnuOptionsDataAccessTypeADO_Click:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmStringFormat, during mnuOptionsDataAccessTypeADO_Click" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_mnuOptionsDataAccessTypeADO_Click
    End Select
    
End Sub
Private Sub mnuOptionsDataAccessTypeDAO_Click()
   
    On Error GoTo Err_mnuOptionsDataAccessTypeDAO_Click
    
    mnuOptionsDataAccessTypeDAO.Checked = Not mnuOptionsDataAccessTypeDAO.Checked
    mnuOptionsDataAccessTypeADO.Checked = Not mnuOptionsDataAccessTypeDAO.Checked
    optADO.Value = Not mnuOptionsDataAccessTypeDAO.Checked
    optDAO.Value = mnuOptionsDataAccessTypeDAO.Checked
    mbUseADO = Not mnuOptionsDataAccessTypeDAO.Checked
    FormatString

Exit_mnuOptionsDataAccessTypeDAO_Click:
    
    On Error GoTo 0
    Exit Sub

Err_mnuOptionsDataAccessTypeDAO_Click:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmStringFormat, during mnuOptionsDataAccessTypeDAO_Click" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_mnuOptionsDataAccessTypeDAO_Click
    End Select
    
End Sub
Private Sub mnuOutputClipboard_Click()
   
    On Error GoTo Err_mnuOutputClipboard_Click
    
    mnuOutputClipboard.Checked = Not mnuOutputClipboard.Checked
    mbClipboard = mnuOutputClipboard.Checked
    mbFile = Not mnuOutputClipboard.Checked
    mnuOutputNotepad.Checked = mbFile
    optClipBoard.Value = mnuOutputClipboard.Checked
    optFile.Value = Not mnuOutputClipboard.Checked

Exit_mnuOutputClipboard_Click:
    
    On Error GoTo 0
    Exit Sub

Err_mnuOutputClipboard_Click:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmStringFormat, during mnuOutputClipboard_Click" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_mnuOutputClipboard_Click
    End Select
    
End Sub
Private Sub mnuOutputNotepad_Click()
   
    On Error GoTo Err_mnuOutputNotepad_Click
    
    mnuOutputNotepad.Checked = Not mnuOutputNotepad.Checked
    mbClipboard = mnuOutputNotepad.Checked
    mbFile = Not mnuOutputNotepad.Checked
    mnuOutputClipboard.Checked = mbFile
    optClipBoard.Value = Not mnuOutputNotepad.Checked
    optFile.Value = mnuOutputNotepad.Checked

Exit_mnuOutputNotepad_Click:
    
    On Error GoTo 0
    Exit Sub

Err_mnuOutputNotepad_Click:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmStringFormat, during mnuOutputNotepad_Click" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_mnuOutputNotepad_Click
    End Select
    
End Sub
Private Sub mnuSelectAgain_Click()
   
    On Error GoTo Err_mnuSelectAgain_Click
    
    SelectAgain

Exit_mnuSelectAgain_Click:
    
    On Error GoTo 0
    Exit Sub

Err_mnuSelectAgain_Click:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmStringFormat, during mnuSelectAgain_Click" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_mnuSelectAgain_Click
    End Select
    
End Sub
Private Sub optADO_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
    On Error GoTo Err_optADO_MouseUp
    
    optDAO.Value = Not optADO.Value
    mbUseADO = optADO.Value
    mnuOptionsDataAccessTypeADO.Checked = optADO.Value
    mnuOptionsDataAccessTypeDAO.Checked = Not optADO.Value
    FormatString

Exit_optADO_MouseUp:
    
    On Error GoTo 0
    Exit Sub

Err_optADO_MouseUp:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmStringFormat, during optADO_MouseUp" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_optADO_MouseUp
    End Select
    
End Sub
Private Sub optClipBoard_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mbClipboard = optClipBoard.Value
    mbFile = Not optClipBoard.Value
    mnuOutputClipboard.Checked = mbClipboard
    mnuOutputNotepad.Checked = mbFile
End Sub
Private Sub optDAO_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
    On Error GoTo Err_optDAO_MouseUp
    
    optADO.Value = Not optDAO.Value
    mbUseADO = optADO.Value
    mnuOptionsDataAccessTypeADO.Checked = Not optDAO.Value
    mnuOptionsDataAccessTypeDAO.Checked = optDAO.Value
    FormatString

Exit_optDAO_MouseUp:
    
    On Error GoTo 0
    Exit Sub

Err_optDAO_MouseUp:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmStringFormat, during optDAO_MouseUp" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_optDAO_MouseUp
    End Select
    
End Sub
Function RemoveChar(sText As String, sChar As String) As String
Attribute RemoveChar.VB_UserMemId = 1610809351
   
    Dim iPos   As Integer
    Dim iStart As Integer
    Dim sTemp  As String
    
    On Error GoTo Err_RemoveChar
    
    iStart = 1
    Do
        iPos = InStr(iStart, sText, sChar)
        If iPos <> 0 Then
            sTemp = sTemp & Mid(sText, iStart, (iPos - iStart))
            iStart = iPos + 1
        End If
    Loop Until iPos = 0
    sTemp = sTemp & Mid(sText, iStart)
    RemoveChar = sTemp

Exit_RemoveChar:
    
    On Error GoTo 0
    Exit Function

Err_RemoveChar:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmStringFormat, during RemoveChar" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_RemoveChar
    End Select
    
End Function
Sub SQLVarPos(ByVal sSQL As String)
Attribute SQLVarPos.VB_UserMemId = 1610809352
   
    Dim iPos As Integer
    Dim iLen As Integer
    
    On Error GoTo Err_SQLVarPos
    
    ReDim aSQLVar(3, 1)
    iLen = 0
    '1 SELECT, INSERT, UPDATE, or DELETE
    iPos = InStr(1, UCase(sSQL), SELECT_STR)
        If iPos = 0 Then
            iPos = InStr(1, UCase(sSQL), INSERT_STR)
                If iPos = 0 Then
                    iPos = InStr(1, UCase(sSQL), UPDATE_STR)
                        If iPos = 0 Then
                            iPos = InStr(1, UCase(sSQL), DELETE_STR)
                        Else
                            iLen = Len(UPDATE_STR)
                        End If
                        If iPos <> 0 Then
                            iLen = Len(DELETE_STR)
                        End If
                Else
                    iLen = Len(INSERT_STR)
                End If
        Else
            iLen = Len(SELECT_STR)
        End If
        If iPos > 0 Then
            aSQLVar(0, 0) = iPos
            aSQLVar(0, 1) = iLen
        Else
            aSQLVar(0, 0) = -1
            aSQLVar(0, 1) = 0
        End If
    '2 FROM Clause
    iPos = InStr(1, UCase(sSQL), FROM_STR)
    iLen = Len(FROM_STR)
        If iPos > 0 Then
            aSQLVar(1, 0) = iPos
            aSQLVar(1, 1) = iLen
        Else
            aSQLVar(1, 0) = -1
            aSQLVar(1, 1) = 0
        End If
    '3 WHERE Clause
    iPos = InStr(1, UCase(sSQL), WHERE_STR)
    iLen = Len(WHERE_STR)
        If iPos > 0 Then
            aSQLVar(2, 0) = iPos
            aSQLVar(2, 1) = iLen
        Else
            aSQLVar(2, 0) = -1
            aSQLVar(2, 1) = 0
        End If
    '4 GROUP BY Clause
    iPos = InStr(1, UCase(sSQL), GROUPBY_STR)
    iLen = Len(GROUPBY_STR)
        If iPos > 0 Then
            aSQLVar(3, 0) = iPos
            aSQLVar(3, 1) = iLen
        Else
            aSQLVar(3, 0) = -1
            aSQLVar(3, 1) = 0
        End If

Exit_SQLVarPos:
    
    On Error GoTo 0
    Exit Sub

Err_SQLVarPos:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmStringFormat, during SQLVarPos" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_SQLVarPos
    End Select
    
End Sub
Function Substitute(szBuff, szOldString, szNewString)
Attribute Substitute.VB_UserMemId = 1610809350
   
    Dim iStart As Long
    Dim iEnd   As Long
    
    On Error GoTo Err_Substitute
    
    ' Find first substring
    iStart = InStr(1, szBuff, szOldString)
    ' Loop through finding substrings
    Do While iStart <> 0
        ' Find end of string
        iEnd = iStart + Len(szOldString)
        ' Concatenate new string
        szBuff = Left(szBuff, iStart - 1) & szNewString & Right(szBuff, Len(szBuff) - iEnd + 1)
        ' Advance past new string
        iStart = iStart + Len(szNewString)
        ' Find next occurrence
        iStart = InStr(iStart, szBuff, szOldString)
    Loop
    Substitute = szBuff

Exit_Substitute:
    
    On Error GoTo 0
    Exit Function

Err_Substitute:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmStringFormat, during Substitute" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_Substitute
    End Select
    
End Function
Private Sub FormatString()
   
    Dim bContinue As Boolean
    Dim sVar      As String
    Dim bQuote    As Boolean
    Dim iCnt      As Integer
    Dim bSQLSmart As Boolean
    
    On Error GoTo Err_FormatString
    
    txtNewString.Text = ""
        If mbUseADO Then
            txtNewString.Text = "'Visual Basic needs you to manually set a reference" & vbNewLine
            txtNewString.Text = txtNewString.Text & "'to Microsoft ActiveX Data Objects (ADO) 2.5 Library (or higher)" & vbNewLine
            txtNewString.Text = txtNewString.Text & "'by going to the menu and selecting it from Project > References" & vbNewLine & vbNewLine
            txtNewString.Text = txtNewString.Text & BuildADOTop & vbNewLine
            txtNewString.Text = txtNewString.Text & BuildSQL(txtOldString.Text) & vbNewLine
            txtNewString.Text = txtNewString.Text & BuildADOBottom
        Else
            txtNewString.Text = "'Visual Basic needs you to manually set a reference" & vbNewLine
            txtNewString.Text = txtNewString.Text & "'to Microsoft Data Access Objects (DAO) 3.5 Library (or higher)" & vbNewLine
            txtNewString.Text = txtNewString.Text & "'by going to the menu and selecting it from Project > References" & vbNewLine & vbNewLine
            txtNewString.Text = txtNewString.Text & BuildDAOTop & vbNewLine
            txtNewString.Text = txtNewString.Text & BuildSQL(txtOldString.Text) & vbNewLine
            txtNewString.Text = txtNewString.Text & BuildDAOBottom
        End If

Exit_FormatString:
    
    On Error GoTo 0
    Exit Sub

Err_FormatString:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmStringFormat, during FormatString" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_FormatString
    End Select

End Sub
Public Sub Finish()
   
    Dim hFile     As Long
    Dim sFilename As String
    Dim iFileName As Integer
    Dim F         As Form
    
    On Error GoTo Err_Finish
    
    If txtNewString.Text = "" Then
        GoTo Exit_Finish
    End If
    If optClipBoard.Value = True Then
        mbClipboard = True
        mbFile = False
    ElseIf optFile.Value = True Then
        mbClipboard = False
        mbFile = True
    End If
    If mbClipboard Then
        ClipboardCopy txtNewString.Text
        MsgBox "Your code is on the clipboard", vbExclamation
    Else
        sFilename = TempFile("vbrsgen")
        'open and save the textbox to a file
        hFile = FreeFile
        Open sFilename For Binary As hFile
        Put #hFile, , txtNewString.Text
        Close hFile
            If Err.Number <> 0 Then
                MsgBox "Problem creating temporary file. The disk may be full or read only", vbExclamation
                Err.Clear
                GoTo Exit_Finish
            End If
        Call Shell("Notepad " & sFilename, vbNormalFocus)
        Kill sFilename
    End If
    
Exit_Finish:
    
    On Error Resume Next
        If Not (F Is Nothing) Then
            Set F = Nothing
        End If
    On Error GoTo 0
    Exit Sub

Err_Finish:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmStringFormat, during Finish" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_Finish
    End Select

End Sub
Private Sub optFile_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
    On Error GoTo Err_optFile_MouseUp
    
    mbFile = optFile.Value
    mbClipboard = Not optFile.Value
    mnuOutputClipboard.Checked = mbClipboard
    mnuOutputNotepad.Checked = mbFile

Exit_optFile_MouseUp:
    
    On Error GoTo 0
    Exit Sub

Err_optFile_MouseUp:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmStringFormat, during optFile_MouseUp" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_optFile_MouseUp
    End Select

End Sub
Private Sub ClearAll()
   
    On Error GoTo Err_ClearAll
    
    If MsgBox("Clear all text boxes?", vbYesNo + vbQuestion, "Clear") = vbYes Then
        txtNewString.Text = ""
        txtOldString.Text = ""
    End If
    Me.Refresh

Exit_ClearAll:
    
    On Error GoTo 0
    Exit Sub

Err_ClearAll:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmStringFormat, during ClearAll" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_ClearAll
    End Select
    
End Sub
Private Sub SelectAgain()
   
    On Error GoTo Err_SelectAgain
    
    On Error Resume Next
    
    txtNewString.Text = ""
    txtOldString.Text = ""
    Unload Me
    Unload frmCriteria
    frmSelectData.Show
    Unload frmTips
    frmSelectData.txtDBPath.Text = gstrDBPath
    frmSelectData.DAOGetObjects
    frmSelectData.lblNoRecords.Caption = "  Please select Tables or Queries  "
    frmSelectData.lblNoRecords.FontBold = True
    frmSelectData.lblNoRecords.FontItalic = True
    
Exit_SelectAgain:
    
    On Error GoTo 0
    Exit Sub

Err_SelectAgain:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmStringFormat, during SelectAgain" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_SelectAgain
    End Select
    
End Sub
