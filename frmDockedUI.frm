VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F74DBB97-4C02-4B5D-AB22-1D7E188F4415}#1.0#0"; "InnovaDSXP.OCX"
Begin VB.Form frmDockedUI 
   ClientHeight    =   10080
   ClientLeft      =   2655
   ClientTop       =   3420
   ClientWidth     =   12270
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDockedUI.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10080
   ScaleWidth      =   12270
   Begin VB.PictureBox picFlood 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7215
      ScaleHeight     =   255
      ScaleWidth      =   3270
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   6765
      Visible         =   0   'False
      Width           =   3270
   End
   Begin InnovaDSXP.DockStudio dsUI 
      Height          =   7635
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9270
      _cx             =   16351
      _cy             =   13467
      Object.Bindings        =   "frmDockedUI.frx":2CFA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      BackColor       =   -2147483633
      EventMode       =   2
      RightToLeft     =   0
      AutoAlignToParent=   0
      Layout          =   "frmDockedUI.frx":2D84
      LastVerbLayoutFilename=   "C:\My Projects\Visual Basic\DHTML Menu Builder 4.8\MainLayout.dsl"
      LanguageFile    =   ""
   End
   Begin MSComctlLib.StatusBar sbDummy 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   9765
      Width           =   12270
      _ExtentX        =   21643
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16660
            MinWidth        =   4410
            Text            =   "Sel Info"
            TextSave        =   "Sel Info"
            Key             =   "sbFlood"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   4419
            MinWidth        =   4410
            Text            =   "Config Info"
            TextSave        =   "Config Info"
            Key             =   "sbConfig"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "frmDockedUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim IsInit As Boolean

Private Sub dsUI_CommandClick(ByVal Command As InnovaDSXP.Command)

    Select Case Command.Name
        Case "mnuFileNewSubEmptyProject"
            frmMain.NewEmptyProject
        Case "mnuFileNewSubFromPreset"
            frmMain.NewFromPreset
        Case "mnuFileNewSubUsingWizard"
            frmMain.NewFromWizard
        Case "mnuFileNewSubFromDir"
            frmMain.NewFromDir
            
        Case "mnuFileOpen"
            frmMain.LoadMenu
            
        Case "mnuFileSave"
            frmMain.SaveMenu False
        Case "mnuFileSaveAs"
            frmMain.FileSaveAs
        Case "mnuFileSaveAsPreset"
            frmMain.SaveAsPreset
        Case "mnuFileSubmitPreset"
            frmMain.SubmitPreset
        Case "mnuFileExportAsHTML"
            frmMain.ExportAsHTML
        Case "mnuFileProjectProperties"
            ProjectPropertiesPage = pppcGeneral
            frmMain.ShowProjectProperties
            
        Case "mnuFileExit"
            Unload Me
            
        Case "mnuEditUndo"
            frmMain.DoUndo
        Case "mnuEditRedo"
            frmMain.DoRedo
        Case "mnuEditFind"
            With dsUI.DockWindows("dwfFind")
                .Visible = True
                frmFind.SwitchToFindMode
                .Size.Height = frmFind.Height + 22 * 15
            End With
            IsFindVisible = True
        Case "mnuEditFindNext"
            frmFind.DoFind
            IsFindVisible = True
        Case "mnuEditReplace"
            With dsUI.DockWindows("dwfFind")
                .Visible = True
                frmFind.SwitchToReplaceMode
                .Size.Height = frmFind.Height + 22 * 15
            End With
            IsFindVisible = True
        Case "mnuEditDelete"
            frmMain.RemoveItem
        Case "mnuEditRename"
            frmMain.RenameItem
        Case "mnuEditPreferences"
            frmMain.ShowPreferences
            
        Case "mnuMenuAddToolbar"
            frmMain.AddToolbar
        Case "mnuMenuAddMenuItem"
            AddMenuItem
            
        Case "mnuToolsInstallMenus"
            Dim dcw As DocumentWindowForm
            
            On Error Resume Next
            Set dcw = dsUI.DocumentWindows("frmInstallMenus")
            
            If dcw Is Nothing Then Set dcw = CreateInstallMenusDoc
            
            If dcw.Visible Then
                dcw.Selected = True
            Else
                dcw.Visible = True
            End If
    End Select
    
    Debug.Print Command.Name

End Sub

Private Sub AddMenuItem()

    If IsTBMapSel Then
        frmMain.MenuAddGroup
    Else
        frmMain.MenuAddCommand
    End If

End Sub

Private Sub dsUI_DockWindowDock(ByVal DockWindow As InnovaDSXP.DockWindow)

    On Error Resume Next

    Dim frm As Form
    Set frm = DockWindow.DockWindows.GetForm(DockWindow.Index).Form

    frm.Width = frm.Width + 15
    frm.Width = frm.Width - 15

End Sub

Private Sub dsUI_DockWindowEnterFocus(ByVal DockWindow As InnovaDSXP.DockWindow)

    On Error Resume Next

    Dim frm As Form
    Set frm = DockWindow.DockWindows.GetForm(DockWindow.Index).Form

    frm.Width = frm.Width + 15
    frm.Width = frm.Width - 15

End Sub

Private Sub dsUI_DockWindowExpand(ByVal DockWindow As InnovaDSXP.DockWindow)

    On Error Resume Next

    Dim frm As Form
    Set frm = DockWindow.DockWindows.GetForm(DockWindow.Index).Form

    frm.Width = frm.Width + 15
    frm.Width = frm.Width - 15

End Sub

Private Sub dsUI_DockWindowFloat(ByVal DockWindow As InnovaDSXP.DockWindow)

    On Error Resume Next

    Dim frm As Form
    Set frm = DockWindow.DockWindows.GetForm(DockWindow.Index).Form

    frm.Width = frm.Width + 15
    frm.Width = frm.Width - 15

End Sub

Private Sub dsUI_DockWindowHide(ByVal DockWindow As InnovaDSXP.DockWindow, ByVal Reason As InnovaDSXP.VisibleStateChangedConstants)

    If DockWindow.Name = "dwfFind" Then IsFindVisible = False

End Sub

Private Sub dsUI_DocumentWindowActivate(ByVal DocumentWindow As InnovaDSXP.DocumentWindow)

    Dim oDocWin As DockWindowForm
    Dim ShowStyleDialogs As Boolean
    Dim lastState As Boolean
    
    If Not IsInit Then Exit Sub
    
    ShowStyleDialogs = (DocumentWindow.Name = "frmMain")

    For Each oDocWin In dsUI.DockWindows
        If Left(oDocWin.Name, 8) = "dwfStyle" Then
            If Not ShowStyleDialogs Then
                oDocWin.Visible = False
            End If
        End If
    Next oDocWin
    
    If ShowStyleDialogs Then frmMain.UpdateControls
    
    Select Case DocumentWindow.Name
        Case "frmInstallMenus"
            dsUI.DockWindows("dwfLivePreview").Visible = False
            lastState = IsFindVisible
            dsUI.DockWindows("dwfFind").Visible = False
            IsFindVisible = lastState
            dsUI.Commands("mnuEdit").Visible = False
            dsUI.Commands("mnuMenu").Visible = False
            dsUI.Commands("mnuToolsApplyStyleFromPreset").Visible = False
            dsUI.CommandBars("MainMenu").Enabled = False
            frmLCMan.InitDlg
            dsUI.CommandBars("MainMenu").Enabled = True
        Case Else
            dsUI.DockWindows("dwfLivePreview").Visible = True
            dsUI.DockWindows("dwfFind").Visible = IsFindVisible
            dsUI.Commands("mnuEdit").Visible = True
            dsUI.Commands("mnuMenu").Visible = True
            dsUI.Commands("mnuToolsApplyStyleFromPreset").Visible = True
    End Select
    
End Sub

Private Sub Form_Load()

    Dim oMyForm As Form
    Dim dcw As DocumentWindowForm
    Dim oDocWin As DockWindowForm
    Dim IsSaved As Boolean
    
    IsInit = False
    
    If Val(GetSetting(App.EXEName, "WinPos", "X")) = 0 Then
        CenterForm Me
    Else
        Left = GetSetting(App.EXEName, "WinPos", "X")
        Top = GetSetting(App.EXEName, "WinPos", "Y")
        Width = GetSetting(App.EXEName, "WinPos", "W")
        Height = GetSetting(App.EXEName, "WinPos", "H")
        
        If Left + Width / 2 > Screen.Width Or Top + Height / 2 > Screen.Height Then
            Left = Screen.Width / 2 - Width / 2
            Top = Screen.Height / 2 - Height / 2
        End If
    End If
    
    If FileExists(App.Path + "\dmbwinconfig.dsl") Then
        IsSaved = True
        dsUI.Layout.LoadFromFile App.Path + "\dmbwinconfig.dsl"
    Else
        IsSaved = False
    End If
    
    Set dcw = dsUI.DocumentWindows.AddForm("frmMain", "Menus Designer", , frmMain, False)
    dcw.WindowState = dsxpDocumentWindowStateMaximized
    dcw.Visible = True
    
    CreateInstallMenusDoc

    LoadForm frmStyleGeneral, "dwfStyleGeneral", "General", 0
    LoadForm frmStyleColor, "dwfStyleColor", "Color", 0, "Color"
    LoadForm frmStyleFont, "dwfStyleFont", "Font", 0, "Font"
    LoadForm frmStyleCursor, "dwfStyleCursor", "Cursor", 0, "Cursor"
    LoadForm frmStyleImage, "dwfStyleImages", "Images", 0, "Image"
    LoadForm frmStyleContainerStyle, "dwfStyleContainerStyle", "Container Style", 0, "Container"
    LoadForm frmStyleContainerSize, "dwfStyleContainerSize", "Container Size", 0, "Size"
    LoadForm frmStyleSelFX, "dwfStyleSelectionEffects", "Selection", 0, "SelEffects"
    LoadForm frmStyleEffects, "dwfStyleEffects", "Effects", 0, "FX"
    
    LoadForm frmLivePreview, "dwfLivePreview", "Preview", 1, , True
    
    LoadForm frmFind, "dwfFind", "Find/Replace", 2
    
    LoadForm frmTBStyleContainerStyle, "dwfStyleTBContainerStyle", "Container Style", 3, "Container"
    LoadForm frmTBStyleContainerSize, "dwfStyleTBContainerSize", "Container Size", 3, "Size"
    
    If Not IsSaved Then
        For Each oDocWin In dsUI.DockWindows
            Select Case oDocWin.Name
                Case "dwfLivePreview"
                    oDocWin.Size.Width = 300 * 15
                Case "dwfFind"
                    oDocWin.Dock dsxpDockWindowPositionInnerTop
                    oDocWin.Visible = False
                Case Else
                    oDocWin.Dock dsxpDockWindowPositionTabbed, "dwfStyleGeneral"
                    oDocWin.Size.Height = 300 * 15
            End Select
        Next oDocWin
    Else
        dsUI.DockWindows("dwfFind").Visible = False
    End If
    
    frmMain.InitDlg
    
    IsInit = True
    
    dsUI_DocumentWindowActivate dsUI.DocumentWindows("frmMain")
    dsUI.DockWindows("dwfStyleGeneral").Activate

End Sub

 Private Function CreateInstallMenusDoc() As DocumentWindow

    Dim dcw As DocumentWindowForm

    ReDim SelSecProjects(0)
    SecProjMode = spmcFromInstallMenus
    Set dcw = dsUI.DocumentWindows.AddForm("frmInstallMenus", "Install Menus", , frmLCMan, False)
    dcw.WindowState = dsxpDocumentWindowStateMaximized
    
    Set CreateInstallMenusDoc = dcw

End Function

Private Sub LoadForm(oMyForm As Form, dockWinName As String, dockWinCaption As String, TypeID As Integer, Optional ImageName As String, Optional IsPreviewWin As Boolean)

    Dim oDocWin As DockWindowForm
    Dim IsSaved As Boolean
    
    On Error Resume Next

    Load oMyForm
    Set oDocWin = dsUI.DockWindows(dockWinName)
    If oDocWin Is Nothing Then
        IsSaved = False
        If ImageName = "" Then
            Set oDocWin = dsUI.DockWindows.AddForm(dockWinName, dockWinCaption, , oMyForm)
        Else
            Set oDocWin = dsUI.DockWindows.AddForm(dockWinName, dockWinCaption, ImageName, oMyForm)
        End If
        oDocWin.Tag = CStr(TypeID)
    Else
        IsSaved = True
        Set oDocWin.Form = oMyForm
    End If
    
    If Not IsSaved Then
        oDocWin.Behaviour.AllowDockTabbed = dsxpValueTrue
        oDocWin.Behaviour.AllowCollapsing = dsxpValueTrue
        oDocWin.DockState = dsxpDockWindowDockStateDocked
        If IsPreviewWin Then
            oDocWin.Dock dsxpDockWindowPositionRight
        Else
            oDocWin.Dock dsxpDockWindowPositionInnerBottom
        End If
        oDocWin.Visible = True
    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    SaveWinPos
    dsUI.Layout.SaveToFile AppPath + "dmbwinconfig.dsl"

End Sub

Private Sub SaveWinPos()

    If WindowState = vbNormal Then
        SaveSetting App.EXEName, "WinPos", "X", Left
        SaveSetting App.EXEName, "WinPos", "Y", Top
        SaveSetting App.EXEName, "WinPos", "W", Width
        SaveSetting App.EXEName, "WinPos", "H", Height
    End If

End Sub

Private Sub Form_Resize()

    Dim i As Integer
    
    On Error Resume Next
    
    If WindowState = vbMinimized Then Exit Sub
    
    For i = 1 To 2
        dsUI.Move 0, 0, Width - GetClientLeft(Me.hWnd), sbDummy.Top
    
        With frmDockedUI.sbDummy
            If WindowState = vbMaximized Then
                .Align = vbAlignNone
                DoEvents
                .Top = Height - .Height - 700
                .Width = Width - 175
                DoEvents
            Else
                .Align = vbAlignBottom
            End If
            .Panels("sbFlood").Width = .Width - 400
            picFlood.Move 30, .Top + 60, .Width - 470, .Height - 75
        End With
        DoEvents
    Next i

End Sub

Private Sub sbDummy_PanelDblClick(ByVal Panel As MSComctlLib.Panel)

    If Panel.key = "sbConfig" Then frmMain.ReSetDefaultConfig

End Sub


