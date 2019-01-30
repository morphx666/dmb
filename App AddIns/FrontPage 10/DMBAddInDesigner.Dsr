VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} DMBAddInDesigner 
   ClientHeight    =   11565
   ClientLeft      =   4335
   ClientTop       =   2235
   ClientWidth     =   6585
   _ExtentX        =   11615
   _ExtentY        =   20399
   _Version        =   393216
   Description     =   "DHTML Menu Builder FrontPage AddIn"
   DisplayName     =   "DHTML Menu Builder"
   AppName         =   "Microsoft FrontPage"
   AppVer          =   "Microsoft FrontPage 10.0"
   LoadName        =   "Startup"
   LoadBehavior    =   3
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Office\FrontPage"
End
Attribute VB_Name = "DMBAddInDesigner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim WithEvents fpApp As FrontPage.Application
Attribute fpApp.VB_VarHelpID = -1
Dim WithEvents fpTBs As Office.CommandBars
Attribute fpTBs.VB_VarHelpID = -1

Dim WithEvents btnDMB As Office.CommandBarButton
Attribute btnDMB.VB_VarHelpID = -1
Dim WithEvents btnProject As Office.CommandBarButton
Attribute btnProject.VB_VarHelpID = -1
Dim WithEvents btnCompile As Office.CommandBarButton
Attribute btnCompile.VB_VarHelpID = -1
Dim WithEvents btnInstall As Office.CommandBarButton
Attribute btnInstall.VB_VarHelpID = -1
Dim dmbTB As Office.CommandBar
Dim ProgID As String

Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub AddInInstance_OnAddInsUpdate(custom() As Variant)
    'Called when the add-in is loaded, unloaded,
    'or another change occurs.
End Sub

Private Sub AddInInstance_OnBeginShutdown(custom() As Variant)
    'Called just before FrontPage begins to shut down.
End Sub

Private Sub AddInInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    
    Set fpApp = Application
    
    ProgID = AddInInst.ProgID
    
    Set fpTBs = fpApp.CommandBars
    For Each dmbTB In fpTBs
        If dmbTB.Name = "DHTML Menu Builder" Then Exit For
    Next dmbTB
    If dmbTB Is Nothing Then
        Set dmbTB = fpTBs.Add(Name:="DHTML Menu Builder", Temporary:=False)
    End If
    
    Set btnDMB = dmbTB.Controls.Add(Type:=msoControlButton, Temporary:=True)
    AddButton btnDMB, "DHTML Menu Builder", "Start DHTML Menu Builder", 126
    
    Set btnProject = dmbTB.Controls.Add(Type:=msoControlButton, Temporary:=True)
    AddButton btnProject, "Project Properties", "Edit this web's menus project", 24
    btnProject.BeginGroup = True
    
    Set btnInstall = dmbTB.Controls.Add(Type:=msoControlButton, Temporary:=True)
    AddButton btnInstall, "Install Menus", "Install the menus", 107
    
    Set btnCompile = dmbTB.Controls.Add(Type:=msoControlButton, Temporary:=True)
    AddButton btnCompile, "Compile", "Generate the menus", 25
    
End Sub

Private Sub AddButton(ByRef btn As CommandBarButton, Title As String, Description As String, IconIdx As Integer)

    With btn
        .Caption = Title
        .DescriptionText = Description
        SetIcon IconIdx
        .PasteFace
        .OnAction = "!<" & ProgID & ">"
        .Enabled = False
    End With

End Sub

Private Sub SetIcon(idx As Integer)

    Set frmRsc.picIcon.Picture = frmRsc.ilIcons.ListImages(idx).Picture
    Clipboard.Clear
    Clipboard.SetData frmRsc.picIcon.Image

End Sub

Private Sub AddInInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    
    Set fpApp = Nothing
    Set fpTBs = Nothing
    Set btnCompile = Nothing
    Set btnDMB = Nothing
    Set btnInstall = Nothing
    Set btnProject = Nothing
    
    Unload frmRsc
    
End Sub

Private Sub AddInInstance_OnStartupComplete(custom() As Variant)
    'Occurs when the startup of the Office host
    'application is complete.
End Sub

Private Sub SetButtonsState(State As Boolean)

    Dim btn As CommandBarButton

    For Each btn In dmbTB.Controls
        btn.Enabled = State
    Next btn

End Sub

Private Sub btnDMB_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)

    Dim dmbPrj As WebFile
    Dim pWin As PageWindowEx
    Dim FileName As String
    
    Set dmbPrj = fpApp.ActiveWeb.LocateFile("menu/Main Web Site 2.0.dmb")
    If dmbPrj Is Nothing Then
        MsgBox "No DMB Project"
    Else
        dmbPrj.Open
    End If
    
    'FileName = GetTEMPPath + "dmbfpaiwp.dmb"
    
    'Set pWin = dmbPrj.Edit(fpPageViewNormal)
    'SaveFile FileName, pWin.Document.DocumentHTML
    'pWin.Close
    
    'Shell GetSetting("DMB", "RegInfo", "InstallPath") + "dmb.exe " + FileName
    
    'RunShellExecute "Open", GetSetting("DMB", "RegInfo", "InstallPath") + "dmb.exe", FileName, 0, 0

End Sub

Private Function GetTEMPPath() As String

    Dim TempPath As String

    TempPath = Environ("TEMP")
    If TempPath = "" Then TempPath = Environ("TMP")
    If TempPath = "" Then TempPath = GetSetting("DMB", "RegInfo", "InstallPath") + "\States\"
    If Right(TempPath, 1) <> "\" Then TempPath = TempPath + "\"
    
    GetTEMPPath = TempPath

End Function

Private Sub fpApp_OnWebClose(ByVal pWeb As FrontPage.Web, Cancel As Boolean)

    SetButtonsState False

End Sub

Private Sub fpApp_OnWebNew(ByVal pWeb As FrontPage.Web)
    
    SetButtonsState True
    
End Sub

Private Sub fpApp_OnWebOpen(ByVal pWeb As FrontPage.Web)
    
    SetButtonsState True

End Sub

Public Function RunShellExecute(sTopic As String, sFile As Variant, sParams As Variant, sDirectory As Variant, nShowCmd As Long) As Long

   Dim hWndDesk As Long
   Dim success As Long
  
  'the desktop will be the
  'default for error messages
   hWndDesk = GetDesktopWindow()
  
  'execute the passed operation
   success = ShellExecute(hWndDesk, sTopic, sFile, sParams, sDirectory, nShowCmd)

  'This is optional. Uncomment the three lines
  'below to have the "Open With.." dialog appear
  'when the ShellExecute API call fails
  'If success = SE_ERR_NOASSOC Then
  '   Call Shell("rundll32.exe shell32.dll,OpenAs_RunDLL " & sFile, vbNormalFocus)
  'End If
  
  RunShellExecute = success
   
End Function
