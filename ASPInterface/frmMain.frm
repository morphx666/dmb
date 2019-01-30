VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DHTML Menu Builder Project Compiler"
   ClientHeight    =   3585
   ClientLeft      =   6345
   ClientTop       =   5415
   ClientWidth     =   5325
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   5325
   Begin VB.PictureBox picItemIcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   180
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   5
      Top             =   1410
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picItemIcon2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   480
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   4
      Top             =   1410
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picRsc 
      Height          =   270
      Left            =   2475
      ScaleHeight     =   210
      ScaleWidth      =   300
      TabIndex        =   3
      Top             =   3135
      Visible         =   0   'False
      Width           =   360
   End
   Begin MSComctlLib.TreeView tvMenus 
      Height          =   1455
      Left            =   1545
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   360
      Visible         =   0   'False
      WhatsThisHelpID =   20000
      Width           =   3090
      _ExtentX        =   5450
      _ExtentY        =   2566
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      Style           =   1
      Appearance      =   1
      OLEDragMode     =   1
      OLEDropMode     =   1
   End
   Begin VB.PictureBox picFlood 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   825
      ScaleHeight     =   270
      ScaleWidth      =   3990
      TabIndex        =   0
      Top             =   2520
      Visible         =   0   'False
      Width           =   3990
   End
   Begin VB.Label lblDummy 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      Height          =   195
      Left            =   1455
      TabIndex        =   2
      Top             =   3210
      Visible         =   0   'False
      Width           =   465
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim IsUpdating As Boolean
Dim FileName As String
Dim ff As Integer

Friend Sub SaveState(dummy As String)

    'dummy function

End Sub

Friend Function GenCmdIcon(dummy As Integer) As Integer

    'dummy function

End Function

Friend Sub UpdateControls()

    'dummy function

End Sub

Private Function LoadMenu(Optional File As String) As Boolean

    Dim sStr As String
    Dim Ans As Integer
    Dim nLines As Integer
    Dim cLine As Integer
        
    On Error GoTo chkError
    
    ff = FreeFile
    Open File For Input As ff
        Do Until EOF(ff) Or sStr = "[RSC]"
            Line Input #ff, sStr
            nLines = nLines + 1
        Loop
    Close ff
    nLines = nLines - 2
    FloodPanel.Caption = "Loading"
    
    Project = GetProjectProperties(File)
    
    Erase MenuGrps: ReDim MenuGrps(0)
    Erase MenuCmds: ReDim MenuCmds(0)
    
    If EOF(ff) Then GoTo ExitSub
    Line Input #ff, sStr
    Do Until EOF(ff) Or sStr = "[RSC]"
        AddMenuGroup Mid$(sStr, 4)
        Do Until EOF(ff) Or sStr = "[RSC]"
            Line Input #ff, sStr
            cLine = cLine + 1: FloodPanel.Value = (cLine / nLines) * 100
            If Left$(sStr, 3) = "[C]" Then
                AddMenuCommand Mid$(sStr, 6), True
            Else
                Exit Do
            End If
        Loop
        If EOF(ff) Then Exit Do
    Loop
    
    LoadMenu = True
ExitSub:
    Close ff
    
    FloodPanel.Value = 0
   
    Exit Function
    
chkError:
    If Err.Number = 53 Or Err.Number = 76 Then
        MsgBox "The project could not be opened because the file " + Project.FileName + " does not exists", vbCritical + vbOKOnly, "Error Opening Project"
    Else
        MsgBox "The project could not be opened. Error (" & Err.Number & ") " + Err.Description, vbCritical + vbOKOnly, "Error Opening Project"
    End If
    Project.HasChanged = False
    GoTo ExitSub

End Function

Private Sub GetPrgPrefs()
    
    #If DEMO = 0 Then
    USER = GetSetting(App.EXEName, "RegInfo", "User", "DEMO")
    COMPANY = GetSetting(App.EXEName, "RegInfo", "Company", "DEMO")
    USERSN = GetSetting(App.EXEName, "RegInfo", "SerialNumber", "")
    #Else
    USER = "DEMO"
    COMPANY = "DEMO"
    USERSN = ""
    #End If
    
    With Preferences
        .AutoRecover = GetSetting(App.EXEName, "Preferences", "AutoRecover", True)
        .OpenLastProject = GetSetting(App.EXEName, "Preferences", "OpenLastProject", True)
        .SepHeight = GetSetting(App.EXEName, "Preferences", "SepHeight", 13)
        .ShowNag = GetSetting(App.EXEName, "Preferences", "ShowNag", True)
        .ShowWarningAddInEditor = GetSetting(App.EXEName, "Preferences", "ShowWarningAIE", True)
        .CommandsInheritance = GetSetting(App.EXEName, "Preferences", "CmdInh", icFirst)
        .GroupsInheritance = GetSetting(App.EXEName, "Preferences", "GrpInh", icFirst)
        .UseLivePreview = GetSetting(App.EXEName, "Preferences", "UseLivePreview", True)
        .DisableUndoRedo = GetSetting(App.EXEName, "Preferences", "DisableUR", False)
        .ImgSpace = Val(GetSetting(App.EXEName, "Preferences", "ImgSpace", 4))
    End With
    
End Sub
