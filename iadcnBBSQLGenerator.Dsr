VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} iadcnVBRSGen 
   ClientHeight    =   8040
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   9630
   _ExtentX        =   16986
   _ExtentY        =   14182
   _Version        =   393216
   Description     =   $"iadcnBBSQLGenerator.dsx":0000
   DisplayName     =   "VBRSGen"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Startup"
   LoadBehavior    =   1
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSafe     =   -1  'True
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "iadcnVBRSGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public FormDisplayed          As Boolean
Public VBInstance             As VBIDE.VBE
Public WithEvents MenuHandler As CommandBarEvents          'command bar event handler
Attribute MenuHandler.VB_VarHelpID = -1

Dim mcbMenuCommandBar         As Office.CommandBarControl
Dim mfrmAddIn                 As New frmSelectData
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
       
    On Error GoTo Err_AddinInstance_OnConnection
    
    'save the VB instance
    Set VBInstance = Application
    If ConnectMode = ext_cm_External Then
        'Used by the wizard toolbar to start this wizard
        Me.Show
    Else
        Set mcbMenuCommandBar = AddToAddInCommandBar("VBRSGen")
        'sink the event
        Set Me.MenuHandler = VBInstance.Events.CommandBarEvents(mcbMenuCommandBar)
    End If
    If ConnectMode = ext_cm_AfterStartup Then
        If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
            'set this to display the form on connect
            Me.Show
        End If
    End If

Exit_AddinInstance_OnConnection:
    
    On Error GoTo 0
    Exit Sub

Err_AddinInstance_OnConnection:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In Connect, during AddinInstance_OnConnection" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_AddinInstance_OnConnection
    End Select
    
End Sub
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    On Error Resume Next
    
    'delete the command bar entry
    mcbMenuCommandBar.Delete
    'shut down the Add-In
    If FormDisplayed Then
        SaveSetting App.Title, "Settings", "DisplayOnConnect", "1"
        FormDisplayed = False
    Else
        SaveSetting App.Title, "Settings", "DisplayOnConnect", "0"
    End If
    Unload mfrmAddIn
    Set mfrmAddIn = Nothing
End Sub
Function AddToAddInCommandBar(sCaption As String) As Office.CommandBarControl

    Dim cbMenuCommandBar As Office.CommandBarControl 'command bar object
    Dim cbMenu           As Object                   'command bar object

    On Error GoTo Err_AddToAddInCommandBar

    'see if we can find the Add-Ins menu
    Set cbMenu = VBInstance.CommandBars("Add-Ins")
        If cbMenu Is Nothing Then
            'not available so we fail
            GoTo Exit_AddToAddInCommandBar
        End If
    'add it to the command bar
    Set cbMenuCommandBar = cbMenu.Controls.Add(1)
    'set the caption
    cbMenuCommandBar.Caption = sCaption
    Set AddToAddInCommandBar = cbMenuCommandBar

Exit_AddToAddInCommandBar:

    On Error Resume Next
        If Not (cbMenuCommandBar Is Nothing) Then
            Set cbMenuCommandBar = Nothing
        End If
        If Not (cbMenu Is Nothing) Then
            Set cbMenu = Nothing
        End If
    On Error GoTo 0
    Exit Function

Err_AddToAddInCommandBar:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In Connect, during AddToAddInCommandBar" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_AddToAddInCommandBar
    End Select

End Function
Sub Hide()
Attribute Hide.VB_UserMemId = 1610809344
   
    On Error GoTo Err_Hide
    
    On Error Resume Next
    
    FormDisplayed = False
    mfrmAddIn.Hide

Exit_Hide:
    
    On Error GoTo 0
    Exit Sub

Err_Hide:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In iadcnVBRSGen, during Hide" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_Hide
    End Select
    
End Sub
Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
   
    On Error GoTo Err_IDTExtensibility_OnStartupComplete
    
    If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
        'set this to display the form on connect
        Me.Show
    End If

Exit_IDTExtensibility_OnStartupComplete:
    
    On Error GoTo 0
    Exit Sub

Err_IDTExtensibility_OnStartupComplete:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In iadcnVBRSGen, during IDTExtensibility_OnStartupComplete" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_IDTExtensibility_OnStartupComplete
    End Select
    
End Sub
Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
   
    'this event fires when the menu is clicked in the IDE
    
    On Error GoTo Err_MenuHandler_Click
    
    Me.Show

Exit_MenuHandler_Click:
    
    On Error GoTo 0
    Exit Sub

Err_MenuHandler_Click:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In iadcnVBRSGen, during MenuHandler_Click" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_MenuHandler_Click
    End Select
    
End Sub
Sub Show()
Attribute Show.VB_UserMemId = 1610809345
   
    On Error GoTo Err_Show
    
    On Error Resume Next
    
    If mfrmAddIn Is Nothing Then
        Set mfrmAddIn = New frmSelectData
    End If
    Set mfrmAddIn.VBInstance = VBInstance
    Set mfrmAddIn.Connect = Me
    FormDisplayed = True
    mfrmAddIn.Show

Exit_Show:
    
    On Error GoTo 0
    Exit Sub

Err_Show:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In iadcnVBRSGen, during Show" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_Show
    End Select
    
End Sub
