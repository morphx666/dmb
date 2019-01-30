VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{9A2947B4-87EE-40EE-A3EF-32BDC32D5726}#1.0#0"; "xfxline3d.ocx"
Object = "{DBF30C82-CAF3-11D5-84FF-0050BA3D926D}#8.5#0"; "VLMnuPlus.ocx"
Object = "{0441225F-21E4-4AB9-94C5-F74C72F390FB}#1.1#0"; "ColorPicker.ocx"
Begin VB.Form frmColorPicker 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Color Picker"
   ClientHeight    =   5925
   ClientLeft      =   6045
   ClientTop       =   5415
   ClientWidth     =   6690
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmColorPicker.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   Begin xfxLine3D.ucLine3D uc3DLine1 
      Height          =   30
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Width           =   4410
      _ExtentX        =   7779
      _ExtentY        =   53
   End
   Begin MSComctlLib.Toolbar tbOptions 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6690
      _ExtentX        =   11800
      _ExtentY        =   953
      ButtonWidth     =   1693
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Used Colors"
            Key             =   "tbUsedColors"
            Style           =   5
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "WebSafe"
            Key             =   "tbWebSafeColor"
            Style           =   1
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picColor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   735
      ScaleHeight     =   240
      ScaleWidth      =   270
      TabIndex        =   0
      Top             =   1590
      Visible         =   0   'False
      Width           =   270
   End
   Begin ColorPicker.ucColorPicker ucCPCtrl 
      Height          =   3750
      Left            =   -30
      TabIndex        =   1
      Top             =   645
      Width           =   4500
      _ExtentX        =   7938
      _ExtentY        =   6615
      ShowWebSafeColor=   -1  'True
   End
   Begin VLMnuPlus.VLMenuPlus vlmCtrl 
      Left            =   5895
      Top             =   4935
      _ExtentX        =   847
      _ExtentY        =   847
      _CXY            =   4
      _CGUID          =   39018.6067476852
      Language        =   0
   End
   Begin VB.Menu mnuColorsMenu 
      Caption         =   "mnuColorsMenu"
      Begin VB.Menu mnuColors 
         Caption         =   "mnuColors"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmColorPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private imColorPopup As New IconMenu6.cIconMenu
Private LastWinState As Integer
Private xMenu As CMenu

Private Sub Form_Load()

    mnuColorsMenu.Visible = False
    
    tbOptions.ImageList = frmMain.ilIcons
    tbOptions.Buttons("tbUsedColors").Image = frmMain.ilIcons.ListImages("mnuMenuColor|mnuContextColor").Index
    tbOptions.Buttons("tbWebSafeColor").Image = frmMain.ilIcons.ListImages("www").Index
    tbOptions.Buttons("tbWebSafeColor").Value = Val(GetSetting(App.EXEName, "ColorPicker", "WebSafe", 1))
    ucCPCtrl.ShowWebSafeColor = (tbOptions.Buttons("tbWebSafeColor").Value = 1)

    Width = ucCPCtrl.Width + Screen.TwipsPerPixelX
    Height = ucCPCtrl.Height + GetClientTop(Me.hwnd) + ucCPCtrl.Top

    CenterForm Me
    SetupCharset Me

    If Not IsDebug Then Set xMenu = New CMenu

    CreateColorPopUp

    ucCPCtrl.Color = SelColor

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    SaveSetting App.EXEName, "ColorPicker", "WebSafe", tbOptions.Buttons("tbWebSafeColor").Value

End Sub

Private Sub mnuColors_Click(Index As Integer)

    If mnuColors(Index).Tag = -2 Then
        ucCPCtrl.Color = -2
    Else
        ucCPCtrl.Color = UsedColors(mnuColors(Index).Tag)
    End If

End Sub

Private Sub tbOptions_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.key
        Case "tbUsedColors"
            DoEvents
            PopupMenu mnuColorsMenu, , tbOptions.Left + Button.Left, tbOptions.Top + tbOptions.Height
        Case "tbWebSafeColor"
            ucCPCtrl.ShowWebSafeColor = (Button.Value = tbrPressed)
    End Select

End Sub

Private Sub tbOptions_ButtonDropDown(ByVal Button As MSComctlLib.Button)

    Select Case Button.key
        Case "tbUsedColors"
            DoEvents
            PopupMenu mnuColorsMenu, , tbOptions.Left + Button.Left, tbOptions.Top + tbOptions.Height
    End Select

End Sub

Private Sub ucCPCtrl_ClickCancel()

    SelColor = -1
    Unload Me

End Sub

Private Sub ucCPCtrl_ClickOK()

    SelColor = ucCPCtrl.Color
    Unload Me

End Sub

Public Sub CreateColorPopUp()

    Dim i As Integer
    Dim ItmName As String
    Dim hc As String
    Dim k As Integer

    Screen.MousePointer = vbArrowHourglass

    k = 0
    For i = 1 To UBound(UsedColors)
        If UsedColors(i) <> -2 Then
            hc = Hex(UsedColors(i))
            ItmName = String(6 - Len(hc), "0") + hc
            ItmName = "#" + Right$(ItmName, 2) + Mid$(ItmName, 3, 2) + Left$(ItmName, 2)
            If k > 0 Then Load mnuColors(k)
            
            ' Need to remove the "#" symbol to prevent Windows with IE7 from crashing
            mnuColors(k).Caption = Replace(ItmName, "#", "")
            mnuColors(k).Tag = i
            k = k + 1
        End If
    Next i

    If SelColor_CanBeTransparent Then
        If k > 0 Then Load mnuColors(k)
        mnuColors(k).Caption = "-"
        Load mnuColors(k + 1)
        mnuColors(k + 1).Caption = "Transparent"
        mnuColors(k + 1).Tag = -2
    End If

    If Not IsDebug Then xMenu.Initialize Me

    If IsWinNT4 Then vlmCtrl.Enabled = False

    Screen.MousePointer = vbDefault

End Sub

Private Sub DrawSelBox()

    picColor.Line (15, 15)-(75, 75), vbWhite, BF
    picColor.Line (30, 30)-(60, 60), vbRed, BF

End Sub

Private Sub ucCPCtrl_UsingPickerEnd()

    Me.WindowState = vbNormal
    frmMain.WindowState = LastWinState

End Sub

Private Sub ucCPCtrl_UsingPickerStart()

    LastWinState = frmMain.WindowState
    Me.WindowState = vbMinimized
    frmMain.WindowState = vbMinimized

End Sub

Private Sub vlmCtrl_SetMenuItemAttributes(ByVal aMenuItem As VLMnuPlus.CMenuItem)

    Dim idx As Integer
    Dim mName As String

    On Error Resume Next

    mName = xMenu.Name(aMenuItem.Caption)

    If InStr(mName, "|") Then
        idx = Val(mnuColors(Split(mName, "|")(1)).Tag)
        With picColor
            .Picture = LoadPicture
            .BackColor = Me.BackColor
            If idx = -2 Then
                frmMain.ilIcons.ListImages("Transparent").Draw .hDc, 0, 0
                If SelColor = -2 Then DrawSelBox
            Else
                picColor.Line (15, 15)-(.Width - 60, .Height - 30), vbBlack, BF
                picColor.Line (30, 30)-(.Width - 75, .Height - 45), UsedColors(idx), BF
                If UsedColors(idx) = SelColor Then DrawSelBox
            End If
            Set .Picture = .Image
            Set aMenuItem.Picture = .Picture
            aMenuItem.CaptionFont.Name = "Courier New"
        End With
    End If

End Sub
