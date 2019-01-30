VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{9A2947B4-87EE-40EE-A3EF-32BDC32D5726}#1.0#0"; "xfxline3d.ocx"
Begin VB.Form frmHotSpotsEditor2 
   Caption         =   "HotSpots Editor 2.0"
   ClientHeight    =   4995
   ClientLeft      =   5310
   ClientTop       =   4755
   ClientWidth     =   7695
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHotSpotsEditor2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   7695
   Begin VB.PictureBox picParams 
      BorderStyle     =   0  'None
      Height          =   4230
      Left            =   4845
      ScaleHeight     =   4230
      ScaleWidth      =   2790
      TabIndex        =   2
      Top             =   45
      Width           =   2790
      Begin VB.CheckBox chkDynaText 
         Caption         =   "Active Text"
         Enabled         =   0   'False
         Height          =   285
         Left            =   1125
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2265
         Width           =   1125
      End
      Begin VB.TextBox txtX 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   0
         TabIndex        =   11
         Text            =   "000"
         Top             =   2265
         Width           =   420
      End
      Begin VB.TextBox txtY 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   585
         TabIndex        =   12
         Text            =   "000"
         Top             =   2265
         WhatsThisHelpID =   20340
         Width           =   420
      End
      Begin VB.PictureBox picHSImage 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   1260
         Left            =   0
         ScaleHeight     =   1230
         ScaleWidth      =   2760
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   2910
         Width           =   2790
      End
      Begin VB.TextBox txtImageName 
         Enabled         =   0   'False
         Height          =   285
         Left            =   0
         TabIndex        =   9
         Top             =   1680
         Width           =   2790
      End
      Begin xfxLine3D.ucLine3D uc3DLine1 
         Height          =   30
         Left            =   0
         TabIndex        =   7
         Top             =   1320
         Width           =   2790
         _ExtentX        =   4921
         _ExtentY        =   53
      End
      Begin MSComctlLib.ImageCombo icmbAlignment 
         Height          =   330
         Left            =   0
         TabIndex        =   6
         Top             =   840
         Width           =   2790
         _ExtentX        =   4921
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
      End
      Begin MSComctlLib.ImageCombo cmbMenuGroups 
         Height          =   330
         Left            =   0
         TabIndex        =   4
         Top             =   210
         Width           =   2790
         _ExtentX        =   4921
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
      End
      Begin xfxLine3D.ucLine3D uc3DLine3 
         Height          =   30
         Left            =   0
         TabIndex        =   15
         Top             =   2715
         Width           =   2790
         _ExtentX        =   4921
         _ExtentY        =   53
      End
      Begin VB.Label lblAttachedGroups 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Attached Group"
         Enabled         =   0   'False
         Height          =   195
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1140
      End
      Begin VB.Label lblAlignment 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alignment"
         Enabled         =   0   'False
         Height          =   195
         Left            =   0
         TabIndex        =   5
         Top             =   630
         Width           =   705
      End
      Begin VB.Label lblPosition 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Position"
         Enabled         =   0   'False
         Height          =   195
         Left            =   0
         TabIndex        =   10
         Top             =   2055
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ","
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   465
         TabIndex        =   14
         Top             =   2340
         Width           =   60
      End
      Begin VB.Label lblImageName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Image Name"
         Enabled         =   0   'False
         Height          =   195
         Left            =   0
         TabIndex        =   8
         Top             =   1455
         Width           =   900
      End
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6465
      TabIndex        =   19
      Top             =   4470
      Width           =   1155
   End
   Begin VB.CommandButton cmdInstall 
      Caption         =   "Install"
      Height          =   360
      Left            =   5205
      TabIndex        =   18
      Top             =   4470
      Width           =   1155
   End
   Begin MSComctlLib.StatusBar sbResizer 
      Align           =   2  'Align Bottom
      Height          =   225
      Left            =   0
      TabIndex        =   20
      Top             =   4770
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   397
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
   End
   Begin xfxLine3D.ucLine3D uc3DLine2 
      Height          =   30
      Left            =   0
      TabIndex        =   17
      Top             =   4365
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   53
   End
   Begin MSComctlLib.ListView lvAnchors 
      Height          =   3750
      Left            =   15
      TabIndex        =   1
      Top             =   255
      Width           =   4725
      _ExtentX        =   8334
      _ExtentY        =   6615
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "chObject"
         Text            =   "Object"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblHS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Available HotSpots"
      Height          =   195
      Left            =   15
      TabIndex        =   0
      Top             =   45
      Width           =   1350
   End
End
Attribute VB_Name = "frmHotSpotsEditor2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type HotSpot
    oCode As String
    Group As String
    Contents As String
    IsImage As Boolean
    ImageName As String
    IsDynaText As Boolean
End Type

Dim HotSpots() As HotSpot
Dim fCode As String
Dim IsUpdating As Boolean
Dim MenuGrps_Back() As MenuGrp
Dim SelId As Integer

Private Sub chkDynaText_Click()

    If IsUpdating Then Exit Sub
    
    HotSpots(lvAnchors.SelectedItem.Index).IsDynaText = (chkDynaText.Value = vbChecked)
    
    lvAnchors_ItemClick lvAnchors.SelectedItem
    lvAnchors.SetFocus

End Sub

Private Sub cmbMenuGroups_Click()

    SelId = cmbMenuGroups.SelectedItem.Index - 1
    UpdateControls

End Sub

Private Sub UpdateControls()

    Dim IsGroupSel As Boolean

    IsUpdating = True
    
    IsGroupSel = cmbMenuGroups.SelectedItem.Index > 1

    With HotSpots(lvAnchors.SelectedItem.Index)
        .Group = IIf(SelId = 0, "", MenuGrps(SelId).Name)
        icmbAlignment.ComboItems(MenuGrps(SelId).Alignment + 1).Selected = True
        txtImageName.Text = IIf(.IsImage, .ImageName, "")
        If .IsDynaText Then .ImageName = txtImageName.Text
        txtX.Text = MenuGrps(SelId).x
        txtY.Text = MenuGrps(SelId).y
        
        txtImageName.Enabled = .IsImage And IsGroupSel
        lblImageName.Enabled = txtImageName.Enabled
        
        chkDynaText.Value = IIf(.IsDynaText, vbChecked, vbUnchecked)
        chkDynaText.Enabled = Not .IsImage And IsGroupSel
        
        txtX.Enabled = Not .IsImage And IsGroupSel And Not .IsDynaText
        txtY.Enabled = txtX.Enabled
        lblPosition.Enabled = Not .IsImage And IsGroupSel
        
        icmbAlignment.Enabled = IsGroupSel And Not .IsDynaText
        If .IsDynaText Then icmbAlignment.ComboItems(1).Selected = True
    End With
    
    lvAnchors.SelectedItem.SmallIcon = IIf(IsGroupSel, HotSpotIcon(HotSpots(lvAnchors.SelectedItem.Index)), IconIndex("EmptyIcon"))
    
    lblAttachedGroups.Enabled = cmbMenuGroups.Enabled
    lblAlignment.Enabled = icmbAlignment.Enabled
    
    IsUpdating = False

End Sub

Private Sub cmdClose_Click()

    MenuGrps = MenuGrps_Back
    HSCanceled = True
    Unload Me

End Sub

Private Sub cmdInstall_Click()

    Dim i As Integer
    Dim moOver As String
    Dim moClick As String
    Dim moDblClick As String
    Dim moOut As String
    Dim mnOver As String
    Dim mnClick As String
    Dim mnDblClick As String
    Dim mnOut As String
    Dim cCode As String
    Dim nCode As String
    Dim sc As String
    Dim hsFile As String
    
    With Project.UserConfigs(Project.DefaultConfig).Frames
        If .UseFrames Then
            For i = 1 To UBound(HotSpots)
                If HotSpots(i).IsImage Then
                    frmHSArrangement.Show vbModal
                    Exit For
                End If
            Next i
        End If
    End With
    
    CompileProject MenuGrps, MenuCmds, Project, Preferences, params, False, True
    
    For i = 1 To UBound(HotSpots)
        With HotSpots(i)
        
            'Get Current event code
            moOver = GetParamVal(.oCode, "onmouseover")
            moClick = GetParamVal(.oCode, "onclick")
            moDblClick = GetParamVal(.oCode, "ondblclick")
            moOut = GetParamVal(.oCode, "onmouseout")
            
            'Remove any code from DMB
            mnOver = RemoveOurEventCode(moOver)
            mnClick = RemoveOurEventCode(moClick)
            mnDblClick = RemoveOurEventCode(moDblClick)
            mnOut = RemoveOurEventCode(moOut)
            
            nCode = .oCode
            If LenB(.Group) <> 0 Then
                SelId = GetIDByName(.Group)
                
                If .IsDynaText Then
                    .ImageName = "dmbHSdyna_" + .Group
                    .IsImage = True
                End If
                
                'Update the image name
                If .IsImage Then
                    If LenB(.ImageName) = 0 Then .ImageName = "agiDMB_" + MenuGrps(SelId).Name
                    MenuGrps(SelId).HSImage = .ImageName
                    If Not .IsDynaText Then
                        If InStr(LCase(.Contents), "name") Then
                            .ImageName = FixItemName(.ImageName)
                            If LCase(.ImageName) = LCase(.Group) Then
                                .ImageName = .ImageName + "_HSImage"
                            End If
                            nCode = ChangeParamVal(nCode, "name", .ImageName)
                        Else
                            nCode = Replace(nCode, "<img ", "<img name=""" + .ImageName + """ ", , , vbTextCompare)
                        End If
                    End If
                End If
            
                'Attach the new code
                cCode = GetGroupEventCode(SelId, IIf(.IsImage, .ImageName, ""), False)
                mnOver = GetEventCode("onmouseover", cCode) + mnOver
                mnClick = GetEventCode("onclick", cCode) + mnClick
                mnDblClick = GetEventCode("ondblclick", cCode) + mnDblClick
                mnOut = GetEventCode("onmouseout", cCode) + mnOut
                
                'Update the hotspot's code
                If LenB(moOut) <> 0 Or LenB(mnOut) <> 0 Then nCode = UpdateEventCode(nCode, moOut, mnOut, "onmouseout")
                If LenB(moDblClick) <> 0 Or LenB(mnDblClick) <> 0 Then nCode = UpdateEventCode(nCode, moDblClick, mnDblClick, "ondblclick")
                If LenB(moClick) <> 0 Or LenB(mnClick) <> 0 Then nCode = UpdateEventCode(nCode, moClick, mnClick, "onclick")
                If LenB(moOver) <> 0 Or LenB(mnOver) <> 0 Then nCode = UpdateEventCode(nCode, moOver, mnOver, "onmouseover")
            Else
                If LenB(moOver) <> 0 Then nCode = Replace(nCode, moOver, mnOver, , , vbTextCompare)
                If LenB(moClick) <> 0 Then nCode = Replace(nCode, moClick, mnClick, , , vbTextCompare)
                If LenB(moDblClick) <> 0 Then nCode = Replace(nCode, moDblClick, mnDblClick, , , vbTextCompare)
                If LenB(moOut) <> 0 Then nCode = Replace(nCode, moOut, mnOut, , , vbTextCompare)
            End If
            
            sc = FindEnclosingChar(nCode, "onmouseover")
            If LenB(sc) <> 0 Then nCode = Replace(nCode, " onmouseover=" + sc + sc, "", , , vbTextCompare)
            sc = FindEnclosingChar(nCode, "onclick")
            If LenB(sc) <> 0 Then nCode = Replace(nCode, " onclick=" + sc + sc, "", , , vbTextCompare)
            sc = FindEnclosingChar(nCode, "ondblclick")
            If LenB(sc) <> 0 Then nCode = Replace(nCode, " ondblclick=" + sc + sc, "", , , vbTextCompare)
            sc = FindEnclosingChar(nCode, "onmouseout")
            If LenB(sc) <> 0 Then nCode = Replace(nCode, " onmouseout=" + sc + sc, "", , , vbTextCompare)
            
            'Update the code in the file
            If .IsDynaText Then
                On Error Resume Next
                FileCopy AppPath + "blank.gif", GetRealLocal.ImagesPath + "blank.gif"
                On Error GoTo 0
                nCode = Replace(nCode, ">" + .Contents + "<", "><img src=""" + GetImagesPath + "blank.gif"" width=1 height=15 border=0 align=middle name=""" + "dmbHSdyna_" + .Group + """>" + .Contents + "<")
            End If
            fCode = Replace(fCode, .oCode, nCode)
        End With
    Next i
    
    CompileProject MenuGrps, MenuCmds, Project, Preferences, params, False, True
    
    hsFile = Project.UserConfigs(Project.DefaultConfig).HotSpotEditor.HotSpotsFile
    fCode = RemoveLoaderCode(fCode, hsFile)
    fCode = AttachLoaderCode(fCode, TrimCR(GenLoaderCode(False, False)))
    
    SaveFile hsFile, fCode

    Unload Me

End Sub

Private Function UpdateEventCode(ByVal srcC As String, ByVal oldC As String, ByVal newC As String, ByVal eventName As String) As String

    Dim sc As String
    
    srcC = Replace(srcC, eventName + " =", eventName + "=")
    srcC = Replace(srcC, eventName + "= ", eventName + "=")

    If LenB(oldC) <> 0 Or LenB(newC) <> 0 Then
        If InStr(LCase(srcC), eventName) Then
            sc = FindEnclosingChar(srcC, eventName)
            oldC = eventName + "=" + sc + oldC + sc
            newC = eventName + "=" + sc + newC + sc
            srcC = Replace(srcC, oldC, newC, , , vbTextCompare)
        Else
            srcC = Left(srcC, 3) + eventName + "=""" + newC + """ " + Mid(srcC, 4)
        End If
    End If
    
    UpdateEventCode = srcC

End Function

Private Function FindEnclosingChar(ByVal dCode As String, eventName As String) As String

    Dim i As Integer
    Dim m As String
    Dim ec As String
    Dim p As Integer
    
    p = InStr(1, dCode, eventName, vbTextCompare)
    
    If p > 0 Then
        dCode = LCase(Mid(dCode, p))
        
        i = InStr(dCode, "=")
        For i = i + 1 To Len(dCode)
            m = Mid(dCode, i, 1)
            If (m = " " Or m = ">" Or m = ";") And LenB(ec) <> 0 Then
                FindEnclosingChar = ec
                Exit Function
            End If
            If (m >= "a" And m <= "z") Or (m >= "0" And m <= "9") Or m = "." Or m = "\" Or m = "/" Then
                FindEnclosingChar = ec
                Exit Function
            Else
                ec = m
            End If
        Next i
    End If
    
End Function

Private Function GetImagesPath() As String

    Dim ThisConfig As ConfigDef
    Dim ImgAbsPath As String

    ThisConfig = Project.UserConfigs(Project.DefaultConfig)

    Select Case ThisConfig.Type
        Case ctcRemote
            ImgAbsPath = SetSlashDir(ThisConfig.RootWeb + ThisConfig.ImagesPath, sdFwd)
            ImgAbsPath = Replace(Mid(ImgAbsPath, InStr(InStr(ImgAbsPath, "//") + 2, ImgAbsPath, "/")), "//", "/")
        Case ctcLocal
            ImgAbsPath = AddFileProtocol(SetSlashDir(ThisConfig.ImagesPath, sdFwd))
        Case ctcCDROM
            ImgAbsPath = "%%REL%%"
    End Select
    
    GetImagesPath = ImgAbsPath

End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF1 Then showHelp "dialogs/hse.htm"

End Sub

Private Sub Form_Load()

    Dim g As Integer
    
    LocalizeUI
    
    HSCanceled = False
    
    caption = "HotSpots Editor 2.0 - [" & EllipseText(frmMain.dummyText, Project.UserConfigs(Project.DefaultConfig).HotSpotEditor.HotSpotsFile, DT_PATH_ELLIPSIS) & "]"
    
    If Val(GetSetting(App.EXEName, "HSEWinPos", "X")) = 0 Then
        CenterForm Me
    Else
        Top = GetSetting(App.EXEName, "HSEWinPos", "X")
        Left = GetSetting(App.EXEName, "HSEWinPos", "Y")
        Width = GetSetting(App.EXEName, "HSEWinPos", "W")
        Height = GetSetting(App.EXEName, "HSEWinPos", "H")
    End If
    
    MenuGrps_Back = MenuGrps
    
    cmbMenuGroups.ImageList = frmMain.ilIcons
    cmbMenuGroups.ComboItems.Add , , "(none)", IconIndex("DisabledGroup")
    For g = 1 To UBound(MenuGrps)
        cmbMenuGroups.ComboItems.Add , , NiceGrpCaption(g), IconIndex("Group")
    Next g
    If cmbMenuGroups.ComboItems.Count Then cmbMenuGroups.ComboItems(1).Selected = True
    
    IsUpdating = True
    BuildAlignmentCombo
    GetHotspots
    IsUpdating = False

End Sub

Private Sub BuildAlignmentCombo()

    Dim nItem As ComboItem
    
    icmbAlignment.ImageList = frmMain.ilAlignment
    
    Set nItem = icmbAlignment.ComboItems.Add(, , GetLocalizedStr(116), 8)
    nItem.tag = 0
    Set nItem = icmbAlignment.ComboItems.Add(, , GetLocalizedStr(117), 2)
    nItem.tag = 1
    Set nItem = icmbAlignment.ComboItems.Add(, , GetLocalizedStr(118), 7)
    nItem.tag = 2
    Set nItem = icmbAlignment.ComboItems.Add(, , GetLocalizedStr(119), 1)
    nItem.tag = 3
    Set nItem = icmbAlignment.ComboItems.Add(, , GetLocalizedStr(120), 4)
    nItem.tag = 4
    Set nItem = icmbAlignment.ComboItems.Add(, , GetLocalizedStr(121), 3)
    nItem.tag = 5
    Set nItem = icmbAlignment.ComboItems.Add(, , GetLocalizedStr(122), 6)
    nItem.tag = 6
    Set nItem = icmbAlignment.ComboItems.Add(, , GetLocalizedStr(123), 5)
    nItem.tag = 7
    Set nItem = icmbAlignment.ComboItems.Add(, , GetLocalizedStr(818), 9)
    nItem.tag = 8
    Set nItem = icmbAlignment.ComboItems.Add(, , GetLocalizedStr(819), 10)
    nItem.tag = 9
    Set nItem = icmbAlignment.ComboItems.Add(, , GetLocalizedStr(820), 11)
    nItem.tag = 10
    Set nItem = icmbAlignment.ComboItems.Add(, , GetLocalizedStr(821), 12)
    nItem.tag = 11
    
    icmbAlignment.ComboItems(1).Selected = True

End Sub

Private Sub GetHotspots()

    Dim atags() As String
    Dim i As Integer
    Dim iconIdx As Integer
    Dim p As Integer
    Dim tmp As String
    
    ReDim atags(0)
    ReDim HotSpots(0)
    
    fCode = LoadFile(GetRealLocal.HotSpotEditor.HotSpotsFile)
    
    fCode = Replace(fCode, "<A ", "<a ")
    fCode = Replace(fCode, "</A>", "</a>")
    fCode = Replace(fCode, "<a" + vbCrLf, "<a ")
    fCode = Replace(fCode, "<A" + vbCrLf, "<a ")
    atags = Split(fCode, "<a ")
    
    If UBound(atags) > 0 Then
    
        lvAnchors.icons = frmMain.ilIcons
        lvAnchors.SmallIcons = frmMain.ilIcons
        
        For i = 1 To UBound(atags)
            p = InStr(atags(i), "</a>")
            If p > 0 Then
                ReDim Preserve HotSpots(UBound(HotSpots) + 1)
                With HotSpots(UBound(HotSpots))
                    .oCode = "<a " + Left(atags(i), p - 1) + "</a>"
                    .Group = GetGroupName(.oCode)
                    .Contents = GetContents(.oCode)
                    .IsImage = (InStr(LCase(.Contents), "<img") > 0)
                    If .IsImage Then
                        .ImageName = GetParamVal(.Contents, "name")
                        .IsDynaText = LCase(Left(.ImageName, 10)) = LCase("dmbHSdyna_")
                        If .IsDynaText Then
                            tmp = RemoveDynaTextCode(.Contents)
                            fCode = Replace(fCode, .Contents, tmp)
                            .oCode = Replace(.oCode, .Contents, tmp)
                            .Contents = tmp
                            .ImageName = ""
                            .IsImage = False
                        End If
                    End If
                    iconIdx = HotSpotIcon(HotSpots(UBound(HotSpots)))
                    lvAnchors.ListItems.Add , , FormatContent(.Contents), iconIdx, iconIdx
                End With
            End If
        Next i
    End If
    
    If lvAnchors.ListItems.Count Then lvAnchors_ItemClick lvAnchors.ListItems(1)
    
    CoolListView lvAnchors

End Sub

Private Function RemoveDynaTextCode(dCode As String) As String

    Dim c As String
    Dim p1 As Long
    Dim p2 As Long
    
    p1 = InStr(dCode, "dmbHSdyna_")
    p2 = InStr(p1, dCode, ">")
    p1 = InStrRev(ccObjectVariableNotSet, "<img", p1, vbTextCompare)
    c = Left(dCode, p1) + Mid(dCode, p2 - p1 + 1)

    RemoveDynaTextCode = c

End Function

Private Function HotSpotIcon(hS As HotSpot) As Integer

    If LenB(hS.Group) = 0 Then
        HotSpotIcon = IconIndex("EmptyIcon")
    Else
        If hS.IsImage And Not hS.IsDynaText Then
            HotSpotIcon = IconIndex("HS-Image")
        Else
            HotSpotIcon = IIf(hS.IsDynaText, IconIndex("HS-DynaText"), IconIndex("HS-Text"))
        End If
    End If
    
End Function

Private Function FormatContent(ByVal c As String) As String

    c = Replace(c, vbCrLf, " ")
    c = Replace(c, vbCr, " ")
    c = Replace(c, vbLf, " ")
    c = Replace(c, vbTab, " ")
    
    FormatContent = c

End Function

Private Function GetContents(c As String) As String

    Dim p1 As Long
    Dim p2 As Long
    
    p1 = InStr(c, ">") + 1
    p2 = InStr(c, "</a>")

    GetContents = Mid(c, p1, p2 - p1)

End Function

Private Function GetGroupName(c As String) As String

    Dim gName As String
    Dim id As Integer

    If InStr(c, "ShowMenu('") > 0 Then
        gName = Split(c, "ShowMenu('")(1)
        gName = Left(gName, InStr(gName, "'") - 1)
        
        id = GetIDByName(gName)
        If id > 0 And id <= UBound(MenuGrps) Then
            If MenuGrps(id).Name = gName Then
                GetGroupName = gName
            End If
        End If
    End If

End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    SaveSetting App.EXEName, "HSEWinPos", "X", Top
    SaveSetting App.EXEName, "HSEWinPos", "Y", Left
    SaveSetting App.EXEName, "HSEWinPos", "W", Width
    SaveSetting App.EXEName, "HSEWinPos", "H", Height

End Sub

Private Sub Form_Resize()

    Dim cTop As Long
    
    cTop = GetClientTop(Me.hwnd)

    If Width < 7815 Then Width = 7815
    If Height < 5400 Then Height = 5400

    picParams.Left = Width - 2970
    lvAnchors.Move 15, 255, Width - 3090, Height - 930 - cTop
    uc3DLine2.Move 0, Height - 560 - cTop, Width + 120, 30
    
    cmdInstall.Move Width - 2610, Height - 455 - cTop
    cmdClose.Move Width - 1350, cmdInstall.Top
    
    CoolListView lvAnchors
    
End Sub

Private Sub icmbAlignment_Click()

    UpdateGrpInfo

End Sub

Private Sub UpdateGrpInfo()

    If IsUpdating Then Exit Sub
    
    With MenuGrps(SelId)
        .Alignment = icmbAlignment.SelectedItem.Index - 1
        If HotSpots(lvAnchors.SelectedItem.Index).IsImage Then HotSpots(lvAnchors.SelectedItem.Index).ImageName = txtImageName.Text
        .x = Val(txtX.Text)
        .y = Val(txtY.Text)
    End With

End Sub

Private Sub lvAnchors_ItemClick(ByVal item As MSComctlLib.ListItem)

    cmbMenuGroups.Enabled = True
    icmbAlignment.Enabled = False
    
    IsUpdating = True
    With HotSpots(lvAnchors.SelectedItem.Index)
        If LenB(.Group) <> 0 Then
            cmbMenuGroups.ComboItems(GetIDByName(.Group) + 1).Selected = True
            txtImageName.Text = .ImageName
        Else
            cmbMenuGroups.ComboItems(1).Selected = True
        End If
        If .IsImage Then
            LoadImage .Contents
        Else
            picHSImage.Picture = LoadPicture()
        End If
    End With
    IsUpdating = False
    
    cmbMenuGroups_Click

End Sub

Private Sub LoadImage(c As String)

    Dim imgSrc As String
    Dim hsFile As String

    imgSrc = GetParamVal(c, "src")
    
    If LenB(imgSrc) <> 0 Then
        hsFile = GetRealLocal.HotSpotEditor.HotSpotsFile
        If Not IsExternalLink(hsFile) Then
            If Left(hsFile, 8) = "file:///" Then
                hsFile = Mid(hsFile, 9)
            Else
                imgSrc = GetFilePath(hsFile) + imgSrc
                imgSrc = SetSlashDir(imgSrc, sdBack)
                imgSrc = RemoveDoubleSlashes(imgSrc)
            End If
            On Error Resume Next
            picHSImage.Picture = LoadPicture(imgSrc)
        End If
    End If
    
End Sub

Private Sub txtY_Change()

    UpdateGrpInfo

End Sub

Private Sub txtY_GotFocus()

    SelAll txtY

End Sub

Private Sub txtImageName_Change()

    UpdateGrpInfo

End Sub

Private Sub txtX_Change()

    UpdateGrpInfo

End Sub

Private Sub txtX_GotFocus()

    SelAll txtX

End Sub

Private Function RemoveOurEventCode(ByVal pCode As String) As String

    Dim cCode As String
    Dim p As Integer
    Dim k As Integer
    
    If InStr(pCode, "function anonymous()") Then
        pCode = Mid$(pCode, InStr(pCode, "{") + 1)
        pCode = Mid$(pCode, 1, InStrRev(pCode, "}") - 1)
    End If

    'Compatibility for older versions of DHTML Menu Builder
    cCode = "if(IE){event.srcElement.style.cursor='"
    If InStr(pCode, cCode) Then
        If InStr(InStr(pCode, cCode), pCode, "';}") Then
            pCode = Left$(pCode, InStr(pCode, cCode) - 1) + Mid$(pCode, InStr(InStr(pCode, cCode), pCode, "';}") + 3)
        Else
            pCode = Left$(pCode, InStr(pCode, cCode) - 1) + Mid$(pCode, InStr(InStr(pCode, cCode), pCode, "'}") + 3)
        End If
    End If
    
    ' Remove any tHideAll's
    cCode = "cFrame.tHideAll();"
    If InStr(pCode, cCode) Then
        pCode = Trim$(TrimCR(Join(Split(pCode, cCode))))
    End If
    cCode = "tHideAll();"
    If InStr(pCode, cCode) Then
        pCode = Trim$(TrimCR(Join(Split(pCode, cCode))))
    End If
    
    ' Remove any Open In New Window Code
    cCode = "javascript:dmbNW=window.open("
    If InStr(pCode, cCode) Then
        pCode = Left$(pCode, InStr(pCode, cCode) - 1) + Mid(pCode, InStr(pCode, "dmbNW.focus();") + 14)
    End If
    
    ' Remove any ShowMenu's
    If InStr(pCode, "cFrame.") Then
        cCode = "cFrame.ShowMenu("
        If InStr(pCode, "ShowMenu(") Then
            k = 1
            For p = InStr(pCode, "ShowMenu(") + 9 To Len(pCode)
                Select Case Mid(pCode, p, 1)
                    Case "(": k = k + 1
                    Case ")": k = k - 1
                End Select
                If k = 0 Then Exit For
            Next p
            If Mid(pCode, p + 1, 1) = ";" Then p = p + 1
            pCode = Left$(pCode, InStr(pCode, cCode) - 1) + Mid$(pCode, p + 1)
        End If
    Else
        'Compatibility for older versions of DHTML Menu Builder
        cCode = "ShowMenu("
        If InStr(pCode, "ShowMenu(") Then
            pCode = Left$(pCode, InStr(pCode, cCode) - 1) + Mid$(pCode, InStr(InStr(pCode, cCode), pCode, "false)") + 6)
        End If
    End If
    
    ' Remove any execURL's
    If InStr(pCode, "execURL(") Then
        cCode = "execURL("
    End If
    If InStr(pCode, "cFrame.execURL(") Then
        cCode = "cFrame.execURL("
    End If
    If InStr(pCode, "execURL(") Then
        pCode = Left$(pCode, InStr(pCode, cCode) - 1) + Mid$(pCode, InStr(InStr(pCode, cCode), pCode, "');") + 4)
    End If
    
    pCode = Replace(pCode, "{;}", ";")
    pCode = Replace(pCode, "{}", "")
    While InStr(pCode, ";;") <> 0
        pCode = Replace(pCode, ";;", ";")
    Wend
    pCode = TrimCR(pCode)
    If pCode = ";" Then pCode = ""
    RemoveOurEventCode = pCode
    
End Function

Private Function GetEventCode(eventName As String, ByVal dCode As String) As String

    Dim EventPos As Integer
    
    If InStr(LCase$(dCode), eventName) Then
        EventPos = InStr(LCase$(dCode), eventName) + Len(eventName) + 2
        dCode = Mid$(dCode, EventPos)
        GetEventCode = Left$(dCode, InStr(dCode, Chr(34)) - 1)
    Else
        GetEventCode = ""
    End If

End Function

Private Function TrimCR(ByVal str As String) As String

    Do While (Left$(str, 1) = vbCr) Or (Left$(str, 1) = vbCrLf) Or (Left$(str, 1) = vbLf)
        str = Mid$(str, 2)
    Loop
    
    Do While (Right$(str, 1) = vbCr) Or (Right$(str, 1) = vbCrLf) Or (Right$(str, 1) = vbLf)
        str = Left$(str, Len(str) - 1)
    Loop
    
    TrimCR = str

End Function

Private Sub LocalizeUI()

    lblHS.caption = GetLocalizedStr(706)
    
    lblAttachedGroups.caption = GetLocalizedStr(707)
    lblAlignment.caption = GetLocalizedStr(115)
    lblImageName.caption = GetLocalizedStr(708)
    lblPosition.caption = GetLocalizedStr(709)
    chkDynaText.caption = GetLocalizedStr(710)
    
    cmdInstall.caption = GetLocalizedStr(468)
    cmdClose.caption = GetLocalizedStr(424)
    
    If Preferences.language <> "eng" Then
        chkDynaText.Width = SetCtrlWidth(chkDynaText)
        cmdInstall.Width = SetCtrlWidth(cmdInstall)
        cmdClose.Width = SetCtrlWidth(cmdClose)
    End If

End Sub
