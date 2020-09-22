VERSION 5.00
Begin VB.Form frmHelp 
   Caption         =   "VBRSGen  -    BB's Recordset and SQL Generator    -     Help"
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   8775
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   8775
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   270
      Left            =   8070
      TabIndex        =   1
      Top             =   5940
      Width           =   675
   End
   Begin VB.TextBox txtHelp 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   5820
      Left            =   105
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   60
      Width           =   8655
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
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
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmHelp, during cmdClose_Click" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_cmdClose_Click
    End Select
    
End Sub
Private Sub Form_Load()
   
    Dim strHelp As String
    
    On Error GoTo Err_Form_Load
    
    strHelp = "VBRSGen will atempt to transform any valid Microsoft Access database table or query into a properly formatted block of code to create a Recordset, including creating and initializing variables for ADO or DAO data access types. Because of the huge variety of possible SQL statements, it's almost impossible for the program to accurately parse every clause exactly as you might expect, so always be sure to test the code and" & Chr(39) & " tweak " & Chr(39) & "it by hand, if necessary, to ensure that it provides the results you need."
    strHelp = strHelp & vbCrLf & vbCrLf & "You can select Microsoft Active Data Objects (ADO) or Data Access Objects (DAO) as your data access model, simply select the proper one from the check box or the Options menu."
    strHelp = strHelp & vbCrLf & vbCrLf & "VBRSGen will create a connection string to the database you select, but if you need to allow users of your program to find a database in a different path, be sure to change the VB code accordingly."
    strHelp = strHelp & vbCrLf & vbCrLf & "VBRSGen can create VB code to open an ODBC connection to your Access database using Jet (DAO) or ADO, but to run the code in your project, you MUST set the appropriate reference(s) in your VB IDE. Select Project > References from the VB menu and click on Microsoft Data Access Objects (DAO) or Microsoft Active Data Objects (ADO), as required. If there selections are not available, you may need to download and install the latest release of Microsoft Data Access Components (MDAC). You can download a file called mdac_typ.exe from http://www.microsoft.com/data/download.htm"
    strHelp = strHelp & vbCrLf & vbCrLf & "To use VBRSGen, select Add-Ins > VBRSGen from the Visual Basic menu. Then on the main screen, select an Access database file (.mdb). Click on one of the options, Tables or Queries, and then choose the object you want to use from the list."
    strHelp = strHelp & "You have 2 choices: (1) If you select a Query, you can then press the Format SQL button to go straight to the SQL generation screen. (2) If you choose a Table, the next thing to do is select the Fields you want to include (hold down Ctrl while clicking with the mouse to select multiple fields). On the next screen you'll see the fields you selected, and if you want to create query criteria, select a field and then a logical operator (ie, =, >=, <>, etc) followed by a value. You may add further criteria by clicking AND or OR and then selecting another field, logical operator, value, etc. "
    strHelp = strHelp & "When you're done creating the criteria, click OK and you'll come to the code generation screen. At the code generation screen, you simply pick any options you want to set, and then press Generate Code to have the VB code created in the box below. Choose the output option you prefer and then press OK to copy the code to the Clipboard or into a text file in Notepad, which you can then paste into your project (or save as a text file)."
    txtHelp.Text = strHelp

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
