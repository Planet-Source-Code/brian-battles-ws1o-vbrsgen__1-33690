VERSION 5.00
Begin VB.Form frmCriteria 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "  BB's SQL Generator - Select Criteria..."
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10125
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
   Icon            =   "frmCriteria.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   10125
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   270
      Left            =   9300
      TabIndex        =   13
      Top             =   300
      Width           =   765
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   270
      Left            =   9300
      TabIndex        =   12
      Top             =   15
      Width           =   765
   End
   Begin VB.Frame fraSelected 
      Caption         =   "  Criteria Selected  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2580
      Left            =   120
      TabIndex        =   28
      Top             =   5520
      Width           =   9900
      Begin VB.TextBox txtCriteria 
         Height          =   2250
         Left            =   75
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   240
         Width           =   9750
      End
   End
   Begin VB.Frame fraFields 
      Caption         =   "  Select Criteria  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   120
      TabIndex        =   20
      Top             =   810
      Width           =   9885
      Begin VB.TextBox txtFieldName 
         Alignment       =   2  'Center
         Height          =   330
         Index           =   0
         Left            =   3180
         Locked          =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   690
         Visible         =   0   'False
         Width           =   2745
      End
      Begin VB.OptionButton optAnd 
         Caption         =   "AND"
         Height          =   210
         Index           =   0
         Left            =   5670
         TabIndex        =   33
         Top             =   1320
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.OptionButton optAnd 
         Caption         =   "AND"
         Height          =   210
         Index           =   1
         Left            =   5670
         TabIndex        =   32
         Top             =   2190
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.OptionButton optAnd 
         Caption         =   "AND"
         Height          =   210
         Index           =   2
         Left            =   5670
         TabIndex        =   31
         Top             =   3045
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.OptionButton optAnd 
         Caption         =   "AND"
         Height          =   210
         Index           =   3
         Left            =   5670
         TabIndex        =   30
         Top             =   3900
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtFieldName 
         Alignment       =   2  'Center
         Height          =   330
         Index           =   4
         Left            =   3180
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   4140
         Visible         =   0   'False
         Width           =   2745
      End
      Begin VB.TextBox txtFieldName 
         Alignment       =   2  'Center
         Height          =   330
         Index           =   3
         Left            =   3210
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   3270
         Visible         =   0   'False
         Width           =   2745
      End
      Begin VB.TextBox txtFieldName 
         Alignment       =   2  'Center
         Height          =   330
         Index           =   2
         Left            =   3195
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   2430
         Visible         =   0   'False
         Width           =   2745
      End
      Begin VB.TextBox txtFieldName 
         Alignment       =   2  'Center
         Height          =   330
         Index           =   1
         Left            =   3180
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1575
         Visible         =   0   'False
         Width           =   2745
      End
      Begin VB.OptionButton optOr 
         Caption         =   "OR"
         Height          =   210
         Index           =   3
         Left            =   6450
         TabIndex        =   25
         Top             =   3900
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.OptionButton optOr 
         Caption         =   "OR"
         Height          =   210
         Index           =   2
         Left            =   6450
         TabIndex        =   24
         Top             =   3045
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.OptionButton optOr 
         Caption         =   "OR"
         Height          =   210
         Index           =   1
         Left            =   6450
         TabIndex        =   23
         Top             =   2190
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.ComboBox cboValues 
         Height          =   330
         Index           =   4
         Left            =   6945
         TabIndex        =   10
         Top             =   4170
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.ComboBox cboOperator 
         Height          =   330
         Index           =   4
         ItemData        =   "frmCriteria.frx":08CA
         Left            =   5970
         List            =   "frmCriteria.frx":08CC
         TabIndex        =   9
         Top             =   4140
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.ComboBox cboValues 
         Height          =   330
         Index           =   3
         Left            =   6960
         TabIndex        =   8
         Top             =   3270
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.ComboBox cboOperator 
         Height          =   330
         Index           =   3
         ItemData        =   "frmCriteria.frx":08CE
         Left            =   5985
         List            =   "frmCriteria.frx":08D0
         TabIndex        =   7
         Top             =   3270
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.ComboBox cboValues 
         Height          =   330
         Index           =   2
         Left            =   6960
         TabIndex        =   6
         Top             =   2430
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.ComboBox cboOperator 
         Height          =   330
         Index           =   2
         ItemData        =   "frmCriteria.frx":08D2
         Left            =   5985
         List            =   "frmCriteria.frx":08D4
         TabIndex        =   5
         Top             =   2430
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.ComboBox cboValues 
         Height          =   330
         Index           =   1
         Left            =   6960
         TabIndex        =   4
         Top             =   1560
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.ComboBox cboOperator 
         Height          =   330
         Index           =   1
         ItemData        =   "frmCriteria.frx":08D6
         Left            =   5985
         List            =   "frmCriteria.frx":08D8
         TabIndex        =   3
         Top             =   1560
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.OptionButton optOr 
         Caption         =   "OR"
         Height          =   210
         Index           =   0
         Left            =   6450
         TabIndex        =   22
         Top             =   1320
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.ComboBox cboValues 
         Height          =   330
         Index           =   0
         Left            =   6945
         TabIndex        =   2
         Top             =   690
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.ComboBox cboOperator 
         Height          =   330
         Index           =   0
         ItemData        =   "frmCriteria.frx":08DA
         Left            =   5970
         List            =   "frmCriteria.frx":08DC
         TabIndex        =   1
         Top             =   690
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.ListBox lstFields 
         Height          =   3840
         ItemData        =   "frmCriteria.frx":08DE
         Left            =   105
         List            =   "frmCriteria.frx":08E0
         TabIndex        =   0
         Top             =   645
         Width           =   3000
      End
      Begin VB.Label lblFieldName 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Field Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   3210
         TabIndex        =   29
         Top             =   390
         Width           =   855
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblValues 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Values"
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
         Left            =   7635
         TabIndex        =   27
         Top             =   390
         Visible         =   0   'False
         Width           =   675
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblOperator 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Comparison"
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
         Left            =   5940
         TabIndex        =   26
         Top             =   390
         Visible         =   0   'False
         Width           =   930
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblFields 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fields"
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
         Left            =   150
         TabIndex        =   21
         Top             =   390
         Width           =   495
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraTableName 
      Caption         =   "  Table Name  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   3405
      TabIndex        =   18
      Top             =   30
      Width           =   3135
      Begin VB.TextBox txtTableName 
         Alignment       =   2  'Center
         Height          =   270
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   225
         Width           =   3000
      End
   End
End
Attribute VB_Name = "frmCriteria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private iLastIndex   As Integer
Private strCriteria  As String
Private strFirstPart As String
Private Sub cboOperator_Change(Index As Integer)
   
    On Error GoTo Err_cboOperator_Change
    
    If cboOperator(Index).ListIndex > -1 Then
        cboValues(Index).Visible = True
        GetValues (Index)
    Else
        cboValues(Index).Visible = False
    End If

Exit_cboOperator_Change:
    
    On Error GoTo 0
    Exit Sub

Err_cboOperator_Change:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmCriteria, during cboOperator_Change" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_cboOperator_Change
    End Select
    
End Sub
Private Sub cboOperator_Click(Index As Integer)
   
    On Error GoTo Err_cboOperator_Click
    
    If cboOperator(Index).ListIndex > -1 Then
        cboValues(Index).Visible = True
        GetValues (Index)
    Else
        cboValues(Index).Visible = False
    End If

Exit_cboOperator_Click:
    
    On Error GoTo 0
    Exit Sub

Err_cboOperator_Click:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmCriteria, during cboOperator_Click" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_cboOperator_Click
    End Select

End Sub
Private Sub cboValues_Click(Index As Integer)
    ' show SQL so far in textbox
    iLastIndex = Index + 1
    txtCriteria.Text = "SELECT " & ShowSQL
    txtCriteria.Text = txtCriteria.Text & ShowCriteria(lstFields.Text, Index)
    lstFields.ListIndex = -1
    optAnd(Index).Visible = True
    optOr(Index).Visible = True
End Sub
Private Sub ClearData()
    Dim C As Control
    
    iLastIndex = 0
    strCriteria = ""
    txtCriteria.Text = ""
    strFirstPart = ""
    lstFields.Clear
    For Each C In Controls
        If TypeOf C Is ComboBox Then
            C.Clear
        ElseIf TypeOf C Is OptionButton Then
            C.Value = False
            C.Visible = False
        End If
    Next
End Sub
Private Sub cmdCancel_Click()
   
    Dim I As Integer
    Dim C As Control
    
    On Error GoTo Err_cmdCancel_Click
    
    ClearData
    Unload Me
    frmSelectData.Show

Exit_cmdCancel_Click:
    
    On Error GoTo 0
    Exit Sub

Err_cmdCancel_Click:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmCriteria, during cmdCancel_Click" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_cmdCancel_Click
    End Select
    
End Sub
Private Sub cmdOK_Click()
   
    On Error GoTo Err_cmdOK_Click
    
    ' use criteria
    SetCriteria
    Unload Me

Exit_cmdOK_Click:
    
    On Error GoTo 0
    Exit Sub

Err_cmdOK_Click:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmCriteria, during cmdOK_Click" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_cmdOK_Click
    End Select
    
End Sub
Private Sub Form_Load()
   
    On Error GoTo Err_Form_Load
    
    LoadForm

Exit_Form_Load:
    
    On Error GoTo 0
    Exit Sub

Err_Form_Load:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmCriteria, during Form_Load" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_Form_Load
    End Select
    
End Sub
Private Sub GetValues(cboValuesIndex As Integer)
   
    Dim strSQL As String
    Dim strCnn As String
    Dim RS     As ADODB.Recordset
    Dim CN     As ADODB.Connection
    Dim I      As Integer
    
    On Error GoTo Err_GetValues
    
    strSQL = "SELECT DISTINCT "
        If lstFields.Text = "" Then
            GoTo Exit_GetValues
        End If
    strSQL = strSQL & lstFields.Text
    strSQL = strSQL & " FROM "
    strSQL = strSQL & txtTableName.Text
    Set CN = New ADODB.Connection
    strCnn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & gstrDBPath & ";Mode=Read;Persist Security Info=False"
    CN.ConnectionString = strCnn
    CN.Open
    Set RS = New ADODB.Recordset
    RS.Open strSQL, CN
        If RS.EOF Then
            cboValues(cboValuesIndex).AddItem "No Records In " & lstFields.Text
        Else
            Do Until RS.EOF
                I = RS.Fields(lstFields.Text).Type
                    ' string values
                    If (I = adVarWChar) Or (I = adChar) Or (I = adBSTR) Or (I = adChar) Or (I = adLongVarChar) Or (I = adLongVarWChar) Or (I = adVarChar) Or (I = adVarWChar) Or (I = adWChar) Then
                        cboValues(cboValuesIndex).AddItem "'" & RS.Fields(lstFields.Text) & "'"
                    ' date/time values
                    ElseIf (I = adDate) Or (I = adDBDate) Or (I = adDBTime) Or (I = adDBTimeStamp) Or (I = adFileTime) Then
                        cboValues(cboValuesIndex).AddItem "#" & RS.Fields(lstFields.Text) & "#"
                    Else
                        ' numeric values
                        cboValues(cboValuesIndex).AddItem RS.Fields(lstFields.Text)
                    End If
                RS.MoveNext
            Loop
        End If
    cboValues(cboValuesIndex).ListIndex = -1

Exit_GetValues:
    
    On Error Resume Next
        If Not (RS Is Nothing) Then
            RS.Close
            Set RS = Nothing
        End If
        If Not (CN Is Nothing) Then
            CN.Close
            Set CN = Nothing
        End If
    On Error GoTo 0
    Exit Sub

Err_GetValues:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmCriteria, during GetValues" & vbCrLf & vbCrLf & "Err.Source: " & Err.Source & vbCrLf & vbCrLf & "Err.LastDllError: " & Err.LastDllError, vbInformation, App.Title & " ADVISORY"
            Resume Exit_GetValues
    End Select
    
End Sub
Private Sub LoadForm()
   
    Dim I As Integer
    
    On Error GoTo Err_LoadForm
        
    For I = 0 To 4
        'If lstFields.Text <> "" Then
        cboOperator(I).AddItem "="
        cboOperator(I).AddItem "<>"
        cboOperator(I).AddItem ">="
        cboOperator(I).AddItem "<="
        cboOperator(I).AddItem "Like"
            'cboOperator(I).Visible = True
        'End If
    Next
    
Exit_LoadForm:
    
    On Error GoTo 0
    Exit Sub

Err_LoadForm:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmCriteria, during LoadForm" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_LoadForm
    End Select
    
End Sub
Private Sub lstFields_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim I As Integer
    
    txtFieldName(iLastIndex).Text = lstFields.Text
    txtFieldName(iLastIndex).Visible = True
    cboOperator(iLastIndex).Visible = True
    cboValues(iLastIndex).Visible = True
    lblValues.Visible = True
    lblFieldName.Visible = True
End Sub
Private Sub optAnd_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   
    On Error GoTo Err_optAnd_MouseUp
    
    If optAnd(Index).Value = True Then
        optOr(Index).Value = False
        cboOperator(Index + 1).Visible = True
        cboValues(Index + 1).Visible = True
        lstFields.SetFocus
    End If

Exit_optAnd_MouseUp:
    
    On Error GoTo 0
    Exit Sub

Err_optAnd_MouseUp:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmCriteria, during optAnd_MouseUp" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_optAnd_MouseUp
    End Select
    
End Sub
Private Sub optOr_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   
    On Error GoTo Err_optOr_MouseUp
    
    If optOr(Index).Value = True Then
        optAnd(Index).Value = False
        cboOperator(Index + 1).Visible = True
        cboValues(Index + 1).Visible = True
    End If

Exit_optOr_MouseUp:
    
    On Error GoTo 0
    Exit Sub

Err_optOr_MouseUp:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmCriteria, during optOr_MouseUp" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_optOr_MouseUp
    End Select
    
End Sub
Private Sub SetCriteria()
   
    On Error GoTo Err_SetCriteria
    
    frmStringFormat.txtOldString.Text = txtCriteria.Text
    frmStringFormat.Show

Exit_SetCriteria:
    
    On Error GoTo 0
    Exit Sub

Err_SetCriteria:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmCriteria, during SetCriteria" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_SetCriteria
    End Select
    
End Sub
Private Function ShowCriteria(strField As String, iIndexOfCombo As Integer) As String

    Dim strSQL As String
    Dim I      As Integer
    
    On Error Resume Next

    I = iIndexOfCombo
        If cboOperator(I).Visible Then
            If strCriteria = "" Then
                strCriteria = " WHERE " & strField & " " & cboOperator(I).Text & " " & cboValues(I).Text
            Else
                If optAnd(I - 1).Value = True Then
                    strCriteria = strCriteria & " AND " & strField & " " & cboOperator(I).Text & " " & cboValues(I).Text
                ElseIf optOr(I - 1).Value = True Then
                    strCriteria = strCriteria & " OR " & strField & " " & cboOperator(I).Text & " " & cboValues(I).Text
                Else
                    ' this shouldn't happen
                End If
            End If
        End If
    strSQL = strSQL & strCriteria
    ShowCriteria = strSQL
End Function
Private Function ShowSQL() As String
    Dim I As Integer
    
    strFirstPart = ""
    For I = 0 To lstFields.ListCount - 1
        If strFirstPart = "" Then
            strFirstPart = lstFields.List(I)
        Else
            strFirstPart = strFirstPart & "," & lstFields.List(I)
        End If
    Next
    strFirstPart = strFirstPart & " FROM " & txtTableName.Text & " "
    ShowSQL = strFirstPart
End Function
