VERSION 5.00
Begin VB.Form frmTips 
   Caption         =   "        BB SQL Generator Tips"
   ClientHeight    =   7035
   ClientLeft      =   2370
   ClientTop       =   2400
   ClientWidth     =   8850
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTips.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   8850
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CheckBox chkLoadTipsAtStartup 
      Caption         =   "&Show Tips at Startup"
      Height          =   210
      Left            =   120
      TabIndex        =   3
      Top             =   6780
      Width           =   1935
   End
   Begin VB.CommandButton cmdNextTip 
      Caption         =   "&Next Tip"
      Height          =   315
      Left            =   8055
      TabIndex        =   2
      Top             =   450
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   6510
      Left            =   120
      Picture         =   "frmTips.frx":030A
      ScaleHeight     =   6450
      ScaleWidth      =   7770
      TabIndex        =   1
      Top             =   120
      Width           =   7830
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "How to use the VBRSGen Recordset and SQL Code Generator..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   510
         TabIndex        =   5
         Top             =   45
         Width           =   5070
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblTipText 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Height          =   5640
         Left            =   135
         TabIndex        =   4
         Top             =   630
         Width           =   7530
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   8055
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmTips"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' in-memory database of tips
Dim Tips       As New Collection

' name of tips file
Const TIP_FILE As String = "TIPOFDAY.TXT"

' Index in collection of tip currently being displayed
Dim CurrentTip As Long
Private Sub DoNextTip()
   
    On Error GoTo Err_DoNextTip
    
    ' Select a tip at random
    
    CurrentTip = Int((Tips.Count * Rnd) + 1)
    ' or cycle through the Tips in order
    '    CurrentTip = CurrentTip + 1
    '    If Tips.Count < CurrentTip Then
    '        CurrentTip = 1
    '    End If
    ' Show it
    frmTips.DisplayCurrentTip

Exit_DoNextTip:
    
    On Error GoTo 0
    Exit Sub

Err_DoNextTip:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmHelp, during DoNextTip" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_DoNextTip
    End Select
   
End Sub
Function LoadTips(sFile As String) As Boolean
   
    Dim NextTip As String   ' Each tip read in from file
    Dim InFile  As Integer   ' Descriptor for file
    
    On Error GoTo Err_LoadTips
    
    ' Obtain the next free file descriptor.
    InFile = FreeFile
    ' Make sure a file is specified.
        If sFile = "" Then
            LoadTips = False
            GoTo Exit_LoadTips
        End If
    ' Make sure the file exists before trying to open it
        If Dir(sFile) = "" Then
            LoadTips = False
            GoTo Exit_LoadTips
        End If
    ' Read the collection from a text file
    Open sFile For Input As InFile
        While Not EOF(InFile)
            Line Input #InFile, NextTip
            Tips.Add NextTip
        Wend
    Close InFile
    ' Display a tip at random.
    DoNextTip
    LoadTips = True

Exit_LoadTips:
    
    On Error GoTo 0
    Exit Function

Err_LoadTips:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmHelp, during LoadTips" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_LoadTips
    End Select
    
End Function
Private Sub chkLoadTipsAtStartup_Click()
   
    ' save whether this form should be displayed at startup
    
    On Error GoTo Err_chkLoadTipsAtStartup_Click
    
    SaveSetting App.EXEName, "Options", "Show Tips at Startup", chkLoadTipsAtStartup.Value

Exit_chkLoadTipsAtStartup_Click:
    
    On Error GoTo 0
    Exit Sub

Err_chkLoadTipsAtStartup_Click:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmHelp, during chkLoadTipsAtStartup_Click" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_chkLoadTipsAtStartup_Click
    End Select
    
End Sub
Private Sub cmdNextTip_Click()
   
    On Error GoTo Err_cmdNextTip_Click
    
    DoNextTip

Exit_cmdNextTip_Click:
    
    On Error GoTo 0
    Exit Sub

Err_cmdNextTip_Click:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmHelp, during cmdNextTip_Click" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_cmdNextTip_Click
    End Select
    
End Sub
Private Sub cmdOK_Click()
   
    On Error GoTo Err_cmdOK_Click
    
    Unload Me

Exit_cmdOK_Click:
    
    On Error GoTo 0
    Exit Sub

Err_cmdOK_Click:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmHelp, during cmdOK_Click" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_cmdOK_Click
    End Select
   
End Sub
Private Sub Form_Load()
   
    On Error GoTo Err_Form_Load
    
    Dim ShowAtStartup As Long
    
    ' See if we should be shown at startup
    ShowAtStartup = GetSetting(App.EXEName, "Options", "Show Tips at Startup", 1)
        If ShowAtStartup = 0 Then
            Unload Me
            GoTo Exit_Form_Load
        End If
    ' Set the checkbox, this will force the value to be written back out to the Registry
    chkLoadTipsAtStartup.Value = vbChecked
    ' Seed Rnd
    Randomize
        ' Read in the tips file and display a tip at random.
        If LoadTips(App.Path & "\" & TIP_FILE) = False Then
            lblTipText.Caption = "That the " & TIP_FILE & " file wasn't found? " & vbCrLf & vbCrLf & "Create a text file named " & TIP_FILE & " using NotePad with 1 tip per line Then place it in the same directory as the application"
        End If

Exit_Form_Load:
    
    On Error GoTo 0
    Exit Sub

Err_Form_Load:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmHelp, during Form_Load" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_Form_Load
    End Select

End Sub
Public Sub DisplayCurrentTip()

    On Error GoTo Err_DisplayCurrentTip
    
    If Tips.Count > 0 Then
        lblTipText.Caption = Tips.Item(CurrentTip)
    End If

Exit_DisplayCurrentTip:
    
    On Error GoTo 0
    Exit Sub

Err_DisplayCurrentTip:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmHelp, during DisplayCurrentTip" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_DisplayCurrentTip
    End Select
    
End Sub
