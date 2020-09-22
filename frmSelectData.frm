VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmSelectData 
   Caption         =   "    VBRSGen BB's VB Recordset Generator  -   Select Data Source"
   ClientHeight    =   8130
   ClientLeft      =   60
   ClientTop       =   585
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
   Icon            =   "frmSelectData.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   10095
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrTime 
      Interval        =   1200
      Left            =   9600
      Top             =   1515
   End
   Begin VB.CommandButton cmdGrid 
      Caption         =   "See Field Values in Grid"
      Height          =   270
      Left            =   5820
      TabIndex        =   16
      Top             =   1635
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Frame fraFields 
      Caption         =   "  Select Fields to Include  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6105
      Left            =   4455
      TabIndex        =   13
      Top             =   1995
      Width           =   5595
      Begin VB.TextBox txtFields 
         Height          =   5340
         Left            =   90
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   675
         Width           =   5445
      End
      Begin VB.CommandButton cmdFmtSQL 
         Caption         =   "Format SQL"
         Height          =   270
         Left            =   1230
         TabIndex        =   7
         Top             =   315
         Width           =   1050
      End
      Begin VB.CommandButton cmdCriteria 
         Caption         =   "Set Criteria"
         Height          =   270
         Left            =   135
         TabIndex        =   6
         Top             =   315
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.ListBox lstFields 
         Height          =   5310
         ItemData        =   "frmSelectData.frx":08CA
         Left            =   90
         List            =   "frmSelectData.frx":08CC
         MultiSelect     =   1  'Simple
         TabIndex        =   5
         Top             =   675
         Width           =   5430
      End
   End
   Begin VB.Frame fraTablesQueries 
      Caption         =   "  Database Tables / Queries  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6105
      Left            =   75
      TabIndex        =   12
      Top             =   1995
      Width           =   4335
      Begin VB.ListBox lstTblsQrys 
         Height          =   5310
         ItemData        =   "frmSelectData.frx":08CE
         Left            =   90
         List            =   "frmSelectData.frx":08D0
         TabIndex        =   4
         Top             =   675
         Width           =   4155
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Refresh"
         Height          =   270
         Left            =   1800
         TabIndex        =   8
         Top             =   315
         Width           =   735
      End
   End
   Begin VB.Frame fraDatabasePath 
      Caption         =   "  Database Path  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   375
      TabIndex        =   11
      Top             =   0
      Width           =   8805
      Begin VB.TextBox txtDBPath 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   90
         TabIndex        =   0
         Top             =   240
         Width           =   7215
      End
      Begin VB.CommandButton cmdGetDatabase 
         Caption         =   "Get Database"
         Height          =   270
         Left            =   7380
         TabIndex        =   1
         Top             =   240
         Width           =   1305
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Cancel"
      Height          =   270
      Left            =   9300
      TabIndex        =   9
      Top             =   15
      Width           =   765
   End
   Begin VB.Frame fraObjectTypes 
      Caption         =   "  Select Data Objects  "
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
      Left            =   1245
      TabIndex        =   10
      Top             =   1020
      Width           =   2100
      Begin VB.OptionButton optQueries 
         Caption         =   "Queries"
         Height          =   210
         Left            =   1125
         TabIndex        =   3
         Top             =   240
         Width           =   930
      End
      Begin VB.OptionButton optTables 
         Caption         =   "Tables"
         Height          =   210
         Left            =   90
         TabIndex        =   2
         Top             =   225
         Width           =   855
      End
   End
   Begin MSComDlg.CommonDialog cdlgDB 
      Left            =   3825
      Top             =   5355
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Select Database"
      Filter          =   "*.mdb,*.*"
      InitDir         =   ".."
   End
   Begin VB.Label lblNoRecords 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Please Select Tables or Queries"
      ForeColor       =   &H000000FF&
      Height          =   885
      Left            =   3420
      TabIndex        =   15
      Top             =   720
      Visible         =   0   'False
      Width           =   6630
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileGetDatabase 
         Caption         =   "Get Database..."
      End
      Begin VB.Menu mnuFileSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileCancel 
         Caption         =   "Cancel"
      End
   End
   Begin VB.Menu mnuDataObjects 
      Caption         =   "Data Objects"
      Begin VB.Menu mnuDataObjectsTables 
         Caption         =   "Tables"
      End
      Begin VB.Menu mnuDataObjectsQueries 
         Caption         =   "Queries"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpHelp 
         Caption         =   "Help"
      End
   End
End
Attribute VB_Name = "frmSelectData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public VBInstance     As VBIDE.VBE
Public Connect        As iadcnVBRSGen

Private mstrSQL       As String
Private mstrTableName As String
Private mbTables      As Boolean
Public Sub ADOGetObjects()
Attribute ADOGetObjects.VB_UserMemId = 1610809345

    Dim cnn        As ADODB.Connection
    Dim rst        As ADODB.Recordset
    Dim strSQL     As String
    Dim strConnect As String

    On Error GoTo Err_ADOGetObjects
    
    If mbTables Then
        ' get the tables
        strSQL = "SELECT "
        strSQL = strSQL & "Name "
        strSQL = strSQL & "FROM "
        strSQL = strSQL & "MSysObjects "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & "("
        strSQL = strSQL & "Type = 1"   ' tables
        strSQL = strSQL & ") "
        strSQL = strSQL & "AND  "
        strSQL = strSQL & "("
        strSQL = strSQL & "Left$(Name,4) <> 'MSys'"
        strSQL = strSQL & ")"
    Else
        ' get the queries
        strSQL = "SELECT "
        strSQL = strSQL & "Name "
        strSQL = strSQL & "FROM "
        strSQL = strSQL & "MSysObjects "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & "("
        strSQL = strSQL & "Type = 5"   ' queries
        strSQL = strSQL & ") "
        strSQL = strSQL & "AND  "
        strSQL = strSQL & "("
        strSQL = strSQL & "Left$(Name,4) <> 'MSys'"
        strSQL = strSQL & ")"
    End If
    Set cnn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    strConnect = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & txtDBPath & ";Persist Security Info=False"
    cnn.Open strConnect
    rst.Open strSQL, cnn
    Do While Not rst.EOF
        lstTblsQrys.AddItem rst.Fields("Name")
        DoEvents
        rst.MoveNext
    Loop
    rst.Close
    cnn.Close

Exit_ADOGetObjects:
    
    On Error Resume Next
        If Not (cnn Is Nothing) Then
            cnn.Close
            Set cnn = Nothing
        End If
        If Not (rst Is Nothing) Then
            rst.Close
            Set rst = Nothing
        End If
    On Error GoTo 0
    Exit Sub

Err_ADOGetObjects:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmSelectData, during ADOGetObjects" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_ADOGetObjects
    End Select
    
End Sub
Private Sub cmdClose_Click()
   
    On Error GoTo Err_cmdClose_Click
    
    On Error Resume Next
    
    Dim F As Form
    
    For Each F In Forms
        Unload F
    Next
    Unload Me

Exit_cmdClose_Click:
    
    On Error GoTo 0
    Exit Sub

Err_cmdClose_Click:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmSelectData, during cmdClose_Click" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_cmdClose_Click
    End Select
    
End Sub
Private Sub cmdCriteria_Click()
   
    Dim I        As Integer
    Dim strQuery As String
    
    On Error GoTo Err_cmdCriteria_Click
    
    strQuery = Replace(lstTblsQrys.Text, Chr(13), " ")
    strQuery = Replace(strQuery, Chr(10), " ")
    strQuery = Replace(strQuery, "  ", " ")
    strQuery = Trim$(strQuery)
    frmCriteria.txtTableName.Text = strQuery
    For I = 0 To lstFields.ListCount - 1
        If lstFields.Selected(I) Then
            frmCriteria.lstFields.AddItem lstFields.List(I)
        End If
    Next
    frmCriteria.Show
    Me.Hide

Exit_cmdCriteria_Click:
    
    On Error GoTo 0
    Exit Sub

Err_cmdCriteria_Click:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmSelectData, during cmdCriteria_Click" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_cmdCriteria_Click
    End Select
    
End Sub
Private Sub cmdFmtSQL_Click()

    Dim strFlds As String
    Dim strFmt  As String
    Dim I       As Integer
    
    On Error GoTo Err_cmdFmtSQL_Click
    
    frmStringFormat.Show
    If optQueries.Value = True Then
        strFmt = Replace(txtFields.Text, Chr(13), " ")
        strFmt = Replace(strFmt, Chr(10), " ")
'        strFmt = Replace(strFmt, "[", "")
'        strFmt = Replace(strFmt, "]", "")
        strFmt = Replace(strFmt, "  ", " ")
        frmStringFormat.txtOldString.Text = ""
        frmStringFormat.txtOldString.Text = strFmt
    Else
        frmStringFormat.txtOldString.Text = ""
        strFlds = ""
        For I = 0 To lstFields.ListCount - 1
            If strFlds = "" Then
                If lstFields.Selected(I) Then
                    strFlds = lstFields.List(I)
                End If
            Else
                If lstFields.Selected(I) Then
                    strFlds = strFlds & ", " & lstFields.List(I)
                End If
            End If
        Next
        strFmt = "SELECT " & strFlds & " FROM " & mstrTableName
        strFmt = Replace(strFmt, Chr(13), " ")
        strFmt = Replace(strFmt, Chr(10), " ")
        strFmt = Replace(strFmt, "[", "")
        strFmt = Replace(strFmt, "]", "")
        strFmt = Replace(strFmt, "  ", " ")
        frmStringFormat.txtOldString.Text = strFmt
        mstrSQL = strFmt
    End If
    Me.Hide

Exit_cmdFmtSQL_Click:
    
    On Error GoTo 0
    Exit Sub

Err_cmdFmtSQL_Click:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmSelectData, during cmdFmtSQL_Click" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_cmdFmtSQL_Click
    End Select
    
End Sub
Private Sub cmdGetDatabase_Click()
   
    On Error GoTo Err_cmdGetDatabase_Click
    
    GetDatabase
    
Exit_cmdGetDatabase_Click:
    
    On Error Resume Next
    lblNoRecords.Caption = "Please Select Tables or Queries"
    lblNoRecords.FontBold = False
    lblNoRecords.FontItalic = False
    Screen.MousePointer = vbDefault
    On Error GoTo 0
    Exit Sub

Err_cmdGetDatabase_Click:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmSelectData, during cmdGetDatabase_Click" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_cmdGetDatabase_Click
    End Select
    
End Sub
Private Sub cmdGrid_Click()
   
    On Error GoTo Err_cmdGrid_Click
    
    If Trim$(lstTblsQrys.Text) <> "" Then
        gstrTableName = Trim$(lstTblsQrys.Text)
        frmDataGrid.Show
    End If

Exit_cmdGrid_Click:
    
    On Error GoTo 0
    Exit Sub

Err_cmdGrid_Click:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmSelectData, during cmdGrid_Click" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_cmdGrid_Click
    End Select
    
End Sub
Private Sub cmdRefresh_Click()
   
    On Error GoTo Err_cmdRefresh_Click
    
    DAOGetObjects

Exit_cmdRefresh_Click:
    
    On Error GoTo 0
    Exit Sub

Err_cmdRefresh_Click:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmSelectData, during cmdRefresh_Click" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_cmdRefresh_Click
    End Select
    
End Sub
Public Sub DAOGetObjects()
   
    Dim dbs As DAO.Database
    Dim tdf As DAO.TableDef
    Dim qdf As DAO.QueryDef

    On Error GoTo Err_DAOGetObjects
    
    Screen.MousePointer = vbHourglass
    lblNoRecords.Caption = "Working, please wait..."
    lblNoRecords.Visible = True
    lblNoRecords.FontBold = True
    lblNoRecords.FontItalic = True
    DoEvents
    If txtDBPath.Text = "" Then
        ' user must have cancelled
        GoTo Exit_DAOGetObjects
    End If
    lstTblsQrys.Clear
    Set dbs = Workspaces(0).OpenDatabase(txtDBPath.Text)
    If optTables.Value = True Then
        fraFields.Caption = "  Select Fields to Include  "
        ' get the tables
        For Each tdf In dbs.TableDefs
            If LCase$(Left$(tdf.Name, 4)) = "msys" Then
                ' skip system tables
            Else
                lstTblsQrys.AddItem tdf.Name
            End If
        Next
    ElseIf optQueries.Value = True Then
        fraFields.Caption = "  SQL Statement from Selected Query  "
        ' get the queries
        For Each qdf In dbs.QueryDefs
            lstTblsQrys.AddItem qdf.Name
        Next
    Else
        fraFields.Caption = "  Select Tables or Queries  "
        lstTblsQrys.AddItem "Select Tables or Queries"
        lstTblsQrys.AddItem "from Select Data Objects"
    End If
       
Exit_DAOGetObjects:
    
    On Error Resume Next
    lblNoRecords.Caption = "Please select Tables or Queries"
    lblNoRecords.Visible = True
    lblNoRecords.FontBold = True
    lblNoRecords.FontItalic = True
        If Not (dbs Is Nothing) Then
            dbs.Close
            Set dbs = Nothing
        End If
        If Not (tdf Is Nothing) Then
            Set tdf = Nothing
        End If
        If Not (qdf Is Nothing) Then
            Set qdf = Nothing
        End If
        If mbTables Then
            lblNoRecords.Caption = "Choose Table"
        Else
            If Me.optQueries.Value = True Then
                lblNoRecords.Caption = "Choose Query"
            Else
                lblNoRecords.Caption = "Choose Table or Query"
            End If
        End If
    lblNoRecords.FontBold = False
    lblNoRecords.FontItalic = False
    Screen.MousePointer = vbDefault
    On Error GoTo 0
    Exit Sub

Err_DAOGetObjects:

    Select Case Err
        Case 0
            Resume Next
        Case 3059 ' cancelled by user
            Resume Exit_DAOGetObjects
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmSelectData, during DAOGetObjects" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_DAOGetObjects
    End Select

End Sub
Private Sub Form_Load()
   
    On Error GoTo Err_Form_Load
    
    If Not gbAlreadyOpen Then
        frmTips.Show vbModal
        gbAlreadyOpen = True
    End If

Exit_Form_Load:
    
    On Error GoTo 0
    Exit Sub

Err_Form_Load:

    Select Case Err
        Case 0, 364 ' object was unloaded
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmSelectData, during Form_Load" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_Form_Load
    End Select
    
End Sub
Private Sub GetDatabase()
   
    On Error GoTo Err_GetDatabase
    
    ' reset all controls' values
    fraFields.Caption = "   Table or Query Fields   "
    fraTablesQueries.Caption = "   Select Table or Query   "
    lstFields.Clear
    lstTblsQrys.Clear
    optQueries.Value = False
    optTables.Value = False
    txtFields.Text = ""
    mbTables = False
    lblNoRecords.Caption = "Please select a database file"
    lblNoRecords.Visible = True
    lblNoRecords.FontBold = True
    lblNoRecords.FontItalic = True
    Screen.MousePointer = vbHourglass
    SelectDatabase
    lblNoRecords.Caption = "Working, please wait..."
    gstrDBPath = Trim$(txtDBPath.Text)
    DAOGetObjects
    cmdRefresh.Enabled = True

Exit_GetDatabase:
    
    On Error GoTo 0
    Exit Sub

Err_GetDatabase:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmSelectData, during GetDatabase" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_GetDatabase
    End Select

End Sub
Private Sub GetFields()

    Dim dbs     As DAO.Database
    Dim tdf     As DAO.TableDef
    Dim qdf     As DAO.QueryDef
    Dim rst     As DAO.Recordset
    Dim I       As Integer
    Dim strFN   As String
    Dim strDesc As String

    On Error GoTo Err_GetFields
    
    cmdCriteria.Visible = False
    cmdFmtSQL.Visible = False
    lblNoRecords.Caption = "Working, please wait..."
    lblNoRecords.FontBold = True
    lblNoRecords.FontItalic = True
    lblNoRecords.Visible = True
    DoEvents
    Screen.MousePointer = vbHourglass
    txtFields.Text = ""
    lstFields.Clear
    Set dbs = Workspaces(0).OpenDatabase(txtDBPath.Text)
    strFN = lstTblsQrys.Text
    lblNoRecords.Caption = ""
    lblNoRecords.FontBold = False
    lblNoRecords.FontItalic = False
    lblNoRecords.Visible = True
    If mbTables Then
        mstrSQL = ""
        Set tdf = dbs.TableDefs(lstTblsQrys.Text)
            For I = 0 To tdf.Properties.Count
                If Format$(tdf.Properties(I).Name) = "Description" Then
                    If Trim$(tdf.Properties(16).Value) <> "" Then  ' description
                        strDesc = tdf.Name & " desc: " & Trim$(tdf.Properties(16).Value) & vbCrLf
                    Else
                        strDesc = ""
                    End If
                Else
                    strDesc = ""
                End If
            Next
            If tdf.Properties(7).Value < 1 Then  ' recordcount
                cmdGrid.Visible = False
                lblNoRecords.Caption = strDesc & "Can't display field values because the selected table" & vbCrLf & "does not contain any records"
                lblNoRecords.Visible = True
            Else
                cmdGrid.Visible = True
                If tdf.RecordCount = 1 Then
                    lblNoRecords.Caption = strDesc & "Selected table contains " & tdf.Properties(7).Value & " record"
                Else
                    lblNoRecords.Caption = strDesc & "Selected table contains " & tdf.Properties(7).Value & " records"
                End If
            End If
            For I = 0 To tdf.Fields.Count - 1
                lstFields.AddItem tdf.Fields(I).Name
                If mstrSQL = "" Then
                    mstrSQL = "SELECT " & tdf.Fields(I).Name
                Else
                    mstrSQL = mstrSQL & ", " & tdf.Fields(I).Name
                End If
                mstrSQL = mstrSQL & " FROM " & tdf.Name
                mstrTableName = tdf.Name
            Next
        cmdCriteria.Visible = True
    Else
        Set qdf = dbs.QueryDefs(lstTblsQrys.Text)
        DoEvents
        txtFields.Text = qdf.SQL
            For I = 0 To qdf.Properties.Count
                If Format$(qdf.Properties(I).Name) = "Description" Then
                    If Trim$(qdf.Properties(I).Value) <> "" Then
                        strDesc = qdf.Name & " desc: " & Trim$(qdf.Properties(I).Value) & vbCrLf
                    Else
                        strDesc = ""
                    End If
                Else
                    strDesc = ""
                End If
            Next
            ' can't display action query results in a grid
            If InStr(1, txtFields.Text, "select into", vbTextCompare) Then
                cmdGrid.Visible = False
                lblNoRecords.Caption = strDesc & "Can't display field values because the query as written" & vbCrLf & "does not return any records"
                lblNoRecords.Visible = True
            ElseIf InStr(1, txtFields.Text, "insert ", vbTextCompare) Then
                cmdGrid.Visible = False
                lblNoRecords.Caption = strDesc & "Can't display field values because the" & vbCrLf & "query as written does not return any records"
                lblNoRecords.Visible = True
            ElseIf InStr(1, txtFields.Text, "update ", vbTextCompare) Then
                cmdGrid.Visible = False
                lblNoRecords.Caption = strDesc & "Can't display field values because the" & vbCrLf & "query as written does not return any records"
                lblNoRecords.Visible = True
            ElseIf InStr(1, txtFields.Text, "delete ", vbTextCompare) Then
                cmdGrid.Visible = False
                lblNoRecords.Caption = strDesc & "Can't display field values because the" & vbCrLf & "query as written does not return any records"
                lblNoRecords.Visible = True
            ElseIf InStr(1, txtFields.Text, "select into", vbTextCompare) Then
                cmdGrid.Visible = False
                lblNoRecords.Caption = strDesc & "Can't display field values because the" & vbCrLf & "query as written does not return any records"
                lblNoRecords.Visible = True
            Else
                Set rst = qdf.OpenRecordset
                If rst.RecordCount < 1 Then
                    lblNoRecords.Caption = strDesc & "Can't display field values because the" & vbCrLf & "query as written does not return any records"
                    lblNoRecords.Visible = True
                    cmdGrid.Visible = False
                Else
                    rst.MoveLast
                    rst.MoveFirst
                    cmdGrid.Visible = True
                    lblNoRecords.Visible = True
                        If rst.RecordCount = 1 Then
                            lblNoRecords.Caption = strDesc & "Selected query will return " & rst.RecordCount & " record"
                        Else
                            lblNoRecords.Caption = strDesc & "Selected query will return " & rst.RecordCount & " records"
                        End If
                End If
            End If
        cmdFmtSQL.Visible = True
    End If
    If lstFields.SelCount > 0 Then
        cmdCriteria.Visible = True
    Else
        cmdCriteria.Visible = False
        cmdFmtSQL.Visible = True
    End If
    
Exit_GetFields:
    
     On Error GoTo 0
        If Not (rst Is Nothing) Then
            rst.Close
            Set rst = Nothing
        End If
        If Not (qdf Is Nothing) Then
            qdf.Close
            Set qdf = Nothing
        End If
        If Not (tdf Is Nothing) Then
            Set tdf = Nothing
        End If
        If Not (dbs Is Nothing) Then
            dbs.Close
            Set dbs = Nothing
        End If
    Screen.MousePointer = vbDefault
    On Error GoTo 0
    Exit Sub

Err_GetFields:

    Select Case Err
        Case 0
            Resume Next
        Case 3265, 3251 ' item not found in this collection; Operation is not supported for this type of object
            cmdCriteria.Visible = False
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmSelectData, during GetFields" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_GetFields
    End Select

End Sub
Private Sub lstFields_Click()
   
    On Error GoTo Err_lstFields_Click
    
    If lstFields.SelCount > 0 Then
        cmdCriteria.Visible = True
    Else
        cmdCriteria.Visible = False
    End If

Exit_lstFields_Click:
    
    On Error GoTo 0
    Exit Sub

Err_lstFields_Click:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmSelectData, during lstFields_Click" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_lstFields_Click
    End Select
    
End Sub
Private Sub lstTblsQrys_Click()
   
    On Error GoTo Err_lstTblsQrys_Click
    
    GetFields

Exit_lstTblsQrys_Click:
    
    On Error GoTo 0
    Exit Sub

Err_lstTblsQrys_Click:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmSelectData, during lstTblsQrys_Click" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_lstTblsQrys_Click
    End Select
    
End Sub
Private Sub optQueries_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
    On Error GoTo Err_optQueries_MouseUp
    
    lblNoRecords.Caption = "Working, please wait..."
    lblNoRecords.FontBold = True
    lblNoRecords.FontItalic = True
    lblNoRecords.Visible = True
    DoEvents
    If optQueries.Value = True Then
        lblNoRecords.Caption = "Getting list of Queries, please wait..."
        DoEvents
        cmdCriteria.Visible = False
        fraFields.Caption = "  SQL Statement from Selected Query  "
        mbTables = False
        optTables.Value = False
        fraTablesQueries.Caption = "Database Queries"
        lstFields.Visible = False
        txtFields.Visible = True
    Else
        lblNoRecords.Caption = "Getting list of Tables, please wait..."
        DoEvents
        fraFields.Caption = "  Select Fields to Include  "
        mbTables = True
        optTables.Value = True
        fraTablesQueries.Caption = "Database Tables"
            If optTables.Value = True Then
                cmdGetDatabase.Visible = True
            Else
                cmdGetDatabase.Visible = False
            End If
        lstFields.Visible = True
        txtFields.Visible = False
    End If
    If Not txtDBPath.Text = "" Then
        DAOGetObjects
    End If

Exit_optQueries_MouseUp:
    
    On Error GoTo 0
    Exit Sub

Err_optQueries_MouseUp:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmSelectData, during optQueries_MouseUp" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_optQueries_MouseUp
    End Select
    
End Sub
Private Sub optTables_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
    On Error GoTo Err_optTables_MouseUp
    
    lblNoRecords.Caption = "Working, please wait..."
    lblNoRecords.FontBold = True
    lblNoRecords.FontItalic = True
    lblNoRecords.Visible = True
    DoEvents
    If optTables.Value = True Then
        lblNoRecords.Caption = "Getting list of Tables, please wait..."
        DoEvents
        fraFields.Caption = "  Select Fields from Selected Table  "
        mbTables = True
        optQueries.Value = False
        fraTablesQueries.Caption = "  Database Tables  "
        lstFields.Visible = True
        txtFields.Visible = False
    Else
        lblNoRecords.Caption = "Getting list of Queries, please wait..."
        DoEvents
        cmdGrid.Visible = False
        mbTables = False
        fraFields.Caption = "  SQL Statement from Selected Query  "
        cmdCriteria.Visible = False
        optQueries.Value = True
        fraTablesQueries.Caption = "  Database Queries  "
            If optQueries.Value = True Then
                cmdGetDatabase.Enabled = True
            Else
                cmdGetDatabase.Enabled = False
            End If
        lstFields.Visible = True
        txtFields.Visible = False
    End If
    If Not txtDBPath.Text = "" Then
        DAOGetObjects
    End If

Exit_optTables_MouseUp:
    
    lblNoRecords.FontBold = False
    lblNoRecords.FontItalic = False
    On Error GoTo 0
    Exit Sub

Err_optTables_MouseUp:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmSelectData, during optTables_MouseUp" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_optTables_MouseUp
    End Select
    
End Sub
Public Sub SelectDatabase()
Attribute SelectDatabase.VB_UserMemId = 1610809344
   
    On Error GoTo Err_SelectDatabase
    
    With cdlgDB
        .CancelError = True ' Causes a trappable error to occur when the user hits the 'Cancel' button
        .DialogTitle = "Select Database"
        .InitDir = App.Path
        .FileName = ""
        .Filter = "All Files (*.*)|*.*|Access Database|*.mdb"
        .FilterIndex = 2
        .flags = cdlOFNCreatePrompt + cdlOFNHideReadOnly + cdlOFNExplorer
        .ShowOpen
            If Err = cdlCancel Then ' 'Cancel' button was hit
                ' add code here when the user hits the 'Cancel' button
                cmdCriteria.Visible = False
                cmdFmtSQL.Visible = False
                cmdGrid.Visible = False
                cmdRefresh.Enabled = False
                lblNoRecords.Caption = "Please press Get Database to proceed"
                lblNoRecords.FontBold = True
                lblNoRecords.Visible = True
                optQueries.Enabled = False
                optTables.Enabled = False
                txtDBPath.Text = ""
                lstFields.Enabled = False
                lstTblsQrys.Enabled = False
                txtFields.Enabled = False
            Else
                cmdCriteria.Visible = False
                cmdFmtSQL.Visible = False
                cmdGrid.Visible = False
                cmdRefresh.Enabled = True
                optQueries.Enabled = True
                optTables.Enabled = True
                lstFields.Enabled = True
                lstTblsQrys.Enabled = True
                txtFields.Enabled = True
            End If
        txtDBPath.Text = .FileName
        End With

Exit_SelectDatabase:
    
    On Error GoTo 0
    Exit Sub

Err_SelectDatabase:

    Select Case Err
        Case 0
            Resume Next
        Case 32755  ' user cancelled
            lstTblsQrys.Clear
            lstFields.Clear
            txtFields.Text = ""
            fraFields.Caption = "  Select Table's Fields or Query's SQL Statement  "
            Resume Exit_SelectDatabase
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmSelectData, during SelectDatabase" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_SelectDatabase
    End Select
    
End Sub
Private Sub tmrTime_Timer()

    On Error GoTo Err_tmrTime_Timer
    
    On Error Resume Next
    
    Me.Caption = " VBRSGen  -  Select Data Source   -    " & Format$(Now(), "DDDD, MMMM d, yyyy   h:nn:ss AMPM")

Exit_tmrTime_Timer:
    
    On Error GoTo 0
    Exit Sub

Err_tmrTime_Timer:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmSelectData, during tmrTime_Timer" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_tmrTime_Timer
    End Select
    
End Sub
