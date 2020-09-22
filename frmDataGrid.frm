VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmDataGrid 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "   BB SQL Generator - Fields in Selected Table"
   ClientHeight    =   8130
   ClientLeft      =   1095
   ClientTop       =   285
   ClientWidth     =   10095
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDataGrid.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   10095
   StartUpPosition =   2  'CenterScreen
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
      TabIndex        =   3
      Top             =   30
      Width           =   3135
      Begin VB.TextBox txtTableName 
         Alignment       =   2  'Center
         Height          =   270
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   225
         Width           =   3000
      End
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   270
      Left            =   9300
      TabIndex        =   2
      Top             =   15
      Width           =   765
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   270
      Left            =   9300
      TabIndex        =   1
      Top             =   300
      Width           =   765
   End
   Begin MSDataGridLib.DataGrid grdDataGrid 
      Bindings        =   "frmDataGrid.frx":08CA
      Height          =   7080
      Left            =   0
      TabIndex        =   0
      Top             =   645
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   12488
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   17
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Selected Fields"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc datPrimaryRS 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   7800
      Visible         =   0   'False
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\VB Projects\DataTest\NWIND.MDB;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\VB Projects\DataTest\NWIND.MDB;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   $"frmDataGrid.frx":08E5
      Caption         =   " "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "frmDataGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strType As String
Private Sub Form_Load()

    On Error GoTo Err_Form_Load
    
    Dim D As DAO.Database
    Dim Q As DAO.QueryDef
        
    Set grdDataGrid.DataSource = Nothing
    txtTableName.Text = gstrTableName
    Set D = Workspaces(0).OpenDatabase(gstrDBPath)
    If strType = "query" Then
        Set Q = D.QueryDefs(gstrTableName)
    ElseIf strType = "table" Then
        ' it's a table
    Else
        ' huh??
    End If
    With datPrimaryRS
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & gstrDBPath & ";Persist Security Info=False"
        If strType = "table" Then
            ' it's a table
            .RecordSource = gstrTableName
        ElseIf strType = "query" Then
            ' it's a query
            .RecordSource = Q.SQL
        Else
            ' huh??
        End If
        .Refresh
    End With
    Set grdDataGrid.DataSource = datPrimaryRS
    grdDataGrid.ReBind

Exit_Form_Load:
    
    On Error Resume Next
        If Not (D Is Nothing) Then
            Set D = Nothing
        End If
        If Not (Q Is Nothing) Then
            Set Q = Nothing
        End If
    On Error GoTo 0
    Exit Sub

Err_Form_Load:

    Select Case Err
        Case 0, 3265, 3604 ' item not found in this collection; Invalid SQL statement
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmDataGrid1, during Form_Load" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_Form_Load
    End Select
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
   
    On Error GoTo Err_Form_Unload
    
    Screen.MousePointer = vbDefault

Exit_Form_Unload:
    
    On Error GoTo 0
    Exit Sub

Err_Form_Unload:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmDataGrid, during Form_Unload" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_Form_Unload
    End Select
    
End Sub
Private Sub datPrimaryRS_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
    ' error handling code
    'If you want to ignore errors, comment out the next line
    'If you want to trap them, add code here to handle them
    MsgBox "Data error event hit Error: " & ErrorNumber & vbCrLf & vbCrLf & Description & vbCrLf & vbCrLf & Scode & vbCrLf & vbCrLf & Source
End Sub
Private Sub cmdClose_Click()
   
    On Error GoTo Err_cmdClose_Click
    
    Unload Me

Exit_cmdClose_Click:
    
    On Error GoTo 0
    Exit Sub

Err_cmdClose_Click:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmDataGrid, during cmdClose_Click" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_cmdClose_Click
    End Select
    
End Sub
