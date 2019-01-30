VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{E1C6DB9D-BD4A-4A61-A759-0CED75D034BF}#43.0#0"; "SmartButton.ocx"
Object = "{9A2947B4-87EE-40EE-A3EF-32BDC32D5726}#1.0#0"; "xfxline3d.ocx"
Begin VB.Form frmRscImages 
   Caption         =   "Images"
   ClientHeight    =   2430
   ClientLeft      =   6795
   ClientTop       =   6015
   ClientWidth     =   5445
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRscImages.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2430
   ScaleWidth      =   5445
   Begin SmartButtonProject.SmartButton cmdSave 
      Height          =   405
      Left            =   1395
      TabIndex        =   11
      Top             =   1290
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   714
      Caption         =   "Save...  "
      Picture         =   "frmRscImages.frx":014A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionLayout   =   5
      PictureLayout   =   3
      OffsetLeft      =   3
      Enabled         =   0   'False
   End
   Begin VB.CommandButton cmdXIcon 
      Caption         =   "Extract Icon..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   45
      TabIndex        =   15
      Top             =   1980
      Width           =   1335
   End
   Begin VB.Timer tmrInit 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   1770
      Top             =   1980
   End
   Begin MSComDlg.CommonDialog cDlgBrowser 
      Left            =   705
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Select an Image File"
      Filter          =   "All Image Files | *.jpg;*.gif"
   End
   Begin SmartButtonProject.SmartButton cmdUpdate 
      Height          =   405
      Left            =   2400
      TabIndex        =   12
      Top             =   1290
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   714
      Caption         =   "Update Location...   "
      Picture         =   "frmRscImages.frx":06E4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionLayout   =   5
      PictureLayout   =   3
      OffsetLeft      =   3
      Enabled         =   0   'False
   End
   Begin SmartButtonProject.SmartButton cmdFromFile 
      Height          =   405
      Left            =   45
      TabIndex        =   10
      Top             =   1290
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   714
      Caption         =   "From File...   "
      Picture         =   "frmRscImages.frx":0A7E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionLayout   =   5
      PictureLayout   =   3
      OffsetLeft      =   3
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3345
      TabIndex        =   17
      Top             =   1980
      Width           =   900
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2325
      TabIndex        =   16
      Top             =   1980
      Width           =   900
   End
   Begin xfxLine3D.ucLine3D uc3DLine1 
      Height          =   30
      Left            =   45
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1845
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   53
   End
   Begin VB.HScrollBar hS 
      Height          =   210
      Left            =   30
      Max             =   1
      Min             =   1
      TabIndex        =   9
      Top             =   945
      Value           =   1
      Width           =   4230
   End
   Begin VB.PictureBox picContainer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   45
      ScaleHeight     =   795
      ScaleWidth      =   4155
      TabIndex        =   0
      Top             =   105
      Width           =   4216
      Begin VB.PictureBox picImage 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   3
         Left            =   3435
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   60
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox picImage 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   2
         Left            =   2370
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   60
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox picImage 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   1
         Left            =   1305
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   60
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox picImage 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   480
         Index           =   0
         Left            =   240
         ScaleHeight     =   450
         ScaleWidth      =   450
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   60
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label lblImageName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ImageName"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   3255
         TabIndex        =   8
         Top             =   570
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lblImageName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ImageName"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   2190
         TabIndex        =   7
         Top             =   570
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lblImageName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ImageName"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   1125
         TabIndex        =   6
         Top             =   570
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lblImageName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ImageName"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   5
         Top             =   570
         Visible         =   0   'False
         Width           =   885
      End
   End
   Begin VB.Label lblS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1290
      TabIndex        =   13
      Top             =   1755
      Visible         =   0   'False
      Width           =   465
   End
End
Attribute VB_Name = "frmRscImages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private SelIndex As Integer
Private Images() As String
Private FileName As String
Private ImageInFocus As Boolean
Private NoRender As Boolean
Private NewSelImage As SelImageDef
Private MaxImages As Integer
Private IsReady As Boolean

Private Sub cmdCancel_Click()

    With SelImage
        Set .Picture = LoadPicture()
        .IsResource = False
        .IsValid = False
        .FileName = vbNullString
    End With

    Unload Me

End Sub

Private Sub cmdFromFile_Click()

    On Error GoTo ExitSub
    
    With SelImage
        Set .Picture = LoadPicture()
        .IsResource = False
        .IsValid = False
        .FileName = vbNullString
    End With

    With cDlgBrowser
        .InitDir = GetRealLocal.ImagesPath
        .DialogTitle = GetLocalizedStr(534)
        If SelImage.SupportsFlash Then
            .filter = SupportedImageFilesFlash
        Else
            If SelImage.LimitToCursors Then
                .filter = SupportedCursorFiles
            Else
                .filter = SupportedImageFiles
            End If
        End If
        .FilterIndex = 0
        .Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly + cdlOFNPathMustExist
        .ShowOpen
    End With
    
    With SelImage
        If IsFlash(cDlgBrowser.FileName) Then
            Set .Picture = LoadResPicture(200, vbResIcon)
        Else
            If IsANI(cDlgBrowser.FileName) Then
                Set .Picture = LoadResPicture(201, vbResIcon)
            Else
                On Error Resume Next
                Set .Picture = OLLoadPicture(cDlgBrowser.FileName)
                On Error GoTo 0
            End If
        End If
        .IsResource = False
        .IsValid = True
        .FileName = cDlgBrowser.FileName
    End With
    
ExitSub:

    If Err.number = 0 Then Unload Me

End Sub

Private Sub GetImages()

    Dim imgCode As String
    Dim ff As Integer
    Dim i As Integer
    Dim j As Integer

    If CurState <> -1 Then
        FileName = UndoStates(CurState).FileName
    Else
        FileName = Project.FileName
    End If
    
    ff = FreeFile
    Open FileName For Binary As #ff
        imgCode = Space(LOF(ff))
        Get #ff, , imgCode
    Close #ff
    Images = Split(Mid(imgCode, InStr(imgCode, "[RSC]") + 5), "RSCImg::")
    
ReStart:
    If SelImage.LimitToCursors Then
        For i = 1 To UBound(Images)
            If Not MatchSpec(CStr(Split(Images(i), "::")(0)), "*.csr;*.cur;*.ani") Then
                For j = i To UBound(Images) - 1
                    Images(j) = Images(j + 1)
                Next j
                ReDim Preserve Images(UBound(Images) - 1)
                GoTo ReStart
            End If
        Next i
    Else
        For i = 1 To UBound(Images)
            If MatchSpec(CStr(Split(Images(i), "::")(0)), "*.csr;*.cur;*.ani") Or Len(CStr(Split(Images(i), "::")(1))) < 6 Then
                For j = i To UBound(Images) - 1
                    Images(j) = Images(j + 1)
                Next j
                ReDim Preserve Images(UBound(Images) - 1)
                GoTo ReStart
            End If
        Next i
    End If
    
    If UBound(Images) > MaxImages Then
        hS.Max = UBound(Images) - (MaxImages - 1)
    Else
        hS.Max = 1
    End If
    
    IsReady = True

End Sub

Private Sub cmdOK_Click()

    SelImage = NewSelImage
    SelImage.IsValid = True
    
    DoEvents
    
    Unload Me

End Sub

Private Sub cmdSave_Click()
    
    On Error GoTo ExitSub
    
    Dim sImage As String
    Dim tImage As String
    Dim idx As Integer
    
    idx = SelIndex - hS.Value
    sImage = picImage(idx).ToolTipText

    With cDlgBrowser
        .InitDir = Project.UserConfigs(Project.DefaultConfig).ImagesPath
        .DialogTitle = GetLocalizedStr(914) + " " + GetFileName(sImage)
        .filter = GetFileName(sImage) + "|" + GetFileName(sImage)
        .FileName = GetFileName(sImage)
        .FilterIndex = 0
        .Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly + cdlOFNPathMustExist
        .ShowSave
    End With
    
    tImage = cDlgBrowser.FileName
    CopyRes2File sImage, tImage
    
ExitSub:

End Sub

Private Sub cmdUpdate_Click()

    On Error GoTo ExitSub
    
    Dim oldImage As String
    Dim newImage As String
    Dim idx As Integer
    Dim i As Integer
    Dim nf As String
    Dim nc As String
    
    idx = SelIndex - hS.Value
    oldImage = LCase(picImage(idx).ToolTipText)
    
    With cDlgBrowser
        .InitDir = Project.UserConfigs(Project.DefaultConfig).ImagesPath
        .DialogTitle = GetLocalizedStr(674) + " " + GetFileName(oldImage)
        .filter = GetFileName(oldImage) + "|" + GetFileName(oldImage) + "|" + SupportedImageFiles
        .FileName = GetFileName(oldImage)
        .FilterIndex = 0
        .Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly + cdlOFNPathMustExist
        .ShowOpen
    End With
    
    newImage = cDlgBrowser.FileName
    
    For i = 1 To UBound(Project.Toolbars)
        With Project.Toolbars(i)
            If LCase(.Image) = oldImage Then .Image = newImage
        End With
    Next i
    
    For i = 1 To UBound(MenuCmds)
        With MenuCmds(i)
            If LCase(.BackImage.NormalImage) = oldImage Then .BackImage.NormalImage = newImage
            If LCase(.BackImage.HoverImage) = oldImage Then .BackImage.HoverImage = newImage
            
            If LCase(.LeftImage.NormalImage) = oldImage Then .LeftImage.NormalImage = newImage
            If LCase(.LeftImage.HoverImage) = oldImage Then .LeftImage.HoverImage = newImage
            If LCase(.RightImage.NormalImage) = oldImage Then .RightImage.NormalImage = newImage
            If LCase(.RightImage.HoverImage) = oldImage Then .RightImage.HoverImage = newImage
        End With
    Next i
    
    For i = 1 To UBound(MenuGrps)
        With MenuGrps(i)
            If LCase(.BackImage.NormalImage) = oldImage Then .BackImage.NormalImage = newImage
            If LCase(.BackImage.HoverImage) = oldImage Then .BackImage.HoverImage = newImage
            
            If LCase(.CornersImages.gcTopLeft) = oldImage Then .CornersImages.gcTopLeft = newImage
            If LCase(.CornersImages.gcTopCenter) = oldImage Then .CornersImages.gcTopCenter = newImage
            If LCase(.CornersImages.gcTopRight) = oldImage Then .CornersImages.gcTopRight = newImage
            
            If LCase(.CornersImages.gcLeft) = oldImage Then .CornersImages.gcLeft = newImage
            If LCase(.CornersImages.gcRight) = oldImage Then .CornersImages.gcRight = newImage
            
            If LCase(.CornersImages.gcBottomLeft) = oldImage Then .CornersImages.gcBottomLeft = newImage
            If LCase(.CornersImages.gcBottomCenter) = oldImage Then .CornersImages.gcBottomCenter = newImage
            If LCase(.CornersImages.gcBottomRight) = oldImage Then .CornersImages.gcBottomRight = newImage
            
            If LCase(.Image) = oldImage Then .Image = newImage
            
            If LCase(.tbiLeftImage.NormalImage) = oldImage Then .tbiLeftImage.NormalImage = newImage
            If LCase(.tbiLeftImage.HoverImage) = oldImage Then .tbiLeftImage.HoverImage = newImage
            If LCase(.tbiRightImage.NormalImage) = oldImage Then .tbiRightImage.NormalImage = newImage
            If LCase(.tbiRightImage.HoverImage) = oldImage Then .tbiRightImage.HoverImage = newImage
            If LCase(.tbiBackImage.NormalImage) = oldImage Then .tbiBackImage.NormalImage = newImage
            If LCase(.tbiBackImage.HoverImage) = oldImage Then .tbiBackImage.HoverImage = newImage
            
            If LCase(.scrolling.DnImage.NormalImage) = oldImage Then .scrolling.DnImage.NormalImage = newImage
            If LCase(.scrolling.DnImage.HoverImage) = oldImage Then .scrolling.DnImage.HoverImage = newImage
            If LCase(.scrolling.UpImage.NormalImage) = oldImage Then .scrolling.UpImage.NormalImage = newImage
            If LCase(.scrolling.UpImage.HoverImage) = oldImage Then .scrolling.UpImage.HoverImage = newImage
        End With
    Next i
    
    For i = 1 To UBound(Images)
        Images(i) = Replace(LCase(Split(Images(i), "::")(0)), LCase(oldImage), newImage) + "::" + Split(Images(i), "::")(1)
    Next i
    
    lblImageName(idx).ForeColor = vbBlack
    lblImageName(idx).caption = GetFileName(newImage)
    picImage(idx).ToolTipText = newImage
    
    With SelImage
        Set .Picture = LoadPicture(cDlgBrowser.FileName)
        .IsResource = False
        .IsValid = True
        .FileName = cDlgBrowser.FileName
    End With
    
    frmMain.SaveState "Update " + GetFileName(oldImage) + " Location"
    
ExitSub:

    If Err.number = 0 Then Unload Me

End Sub

Private Sub cmdXIcon_Click()
    
    frmXIcon.Show vbModal

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF1 Then showHelp "dialogs/images_browser.htm"

End Sub

Private Sub Form_Load()

    CenterForm Me
    LocalizeUI
    
    IsReady = False
    
    MaxImages = 4
    
    hS.min = 1
    hS.Value = 1
    hS.Max = 1
    
    cmdUpdate.Enabled = False
    
    tmrInit.Enabled = True
    
End Sub

Private Sub InitDialog()

    Dim i As Integer
    
    Me.Enabled = False
    Screen.MousePointer = vbHourglass
    
    For i = 0 To MaxImages - 1
        picImage(i).Visible = False
        lblImageName(i).Visible = False
    Next i
    
    GetImages
    
    SelIndex = -1
    NoRender = True
    GetImagesFromResource
    
    For i = hS.min To hS.Max
        hS.Value = i
        If ImageInFocus Then
            cmdUpdate.Enabled = True
            Exit For
        End If
    Next i
    
    If (i - 1) = hS.Max And Not ImageInFocus Then
        hS.Value = hS.min
    Else
        For i = i To SelIndex
            If i < hS.Max Then hS.Value = i
        Next i
    End If
    
    NoRender = False
    If SelIndex = -1 Then SelIndex = 1
    GetImagesFromResource
    
    Screen.MousePointer = vbDefault
    Me.Enabled = True
    
    picImage_Click SelIndex - hS.Value

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    SelImage.SupportsFlash = False
    SelImage.LimitToCursors = False

End Sub

Private Sub Form_Resize()

    Dim i As Integer
    Dim cx As Long
    Dim cy As Long
    Dim nH As Integer
    Dim nV As Integer
    Static IsResizing As Boolean
    Dim cTop As Long
    
    If IsResizing Or Not IsReady Then Exit Sub
    IsResizing = True

    If Width < 4455 Then Width = 4455
    If Height < 2835 Then Height = 2835
    
    cTop = GetClientTop(Me.hwnd)

    picContainer.Move 45, 105, Width - 239, Height - 1500 - cTop
    uc3DLine1.Move 45, Height - 575 - cTop, Width - 240, 30
    hS.Move 30, Height - 1890, Width - 225, 210
    cmdOK.Move Width - 2160, Height - 450 - cTop
    cmdCancel.Move Width - 1110, Height - 450 - cTop
    cmdXIcon.Move 45, Height - 450 - cTop
    
    cmdFromFile.Top = Height - 1545
    cmdSave.Top = cmdFromFile.Top
    cmdUpdate.Top = cmdFromFile.Top
    
    nH = ((picContainer.Width / 480) * 480) / 1054
    nV = ((picContainer.Height / 480) * 480) / 855
    If nV * nH <= UBound(Images) Then
        MaxImages = nH * nV
        hS.Max = UBound(Images) - (MaxImages - 1)
    Else
        MaxImages = UBound(Images)
        hS.Value = 1
        hS.Max = 1
    End If
    
    If MaxImages > 0 Then
        hS.SmallChange = 1
        hS.LargeChange = MaxImages
    End If
    
    On Error Resume Next
    For i = picImage.Count - 1 To MaxImages - 1
        Load picImage(i)
        Load lblImageName(i)
    Next i
    
    cx = 240
    cy = 60
    For i = 0 To MaxImages - 1
        picImage(i).Move cx, cy
        lblImageName(i).Move cx - 180, cy + 510
        
        cx = cx + 1065
        If cx + picImage(i).Width + 240 > picContainer.Width Then
            cx = 240
            cy = cy + 800
        End If
    Next i
    
    GetImagesFromResource
    
    IsResizing = False

End Sub

Private Sub hS_Change()

    GetImagesFromResource

End Sub

Private Sub hS_Scroll()

    GetImagesFromResource

End Sub

Private Sub picImage_Click(Index As Integer)

    Dim i As Integer
    
    For i = 0 To MaxImages - 1
        picImage(i).BorderStyle = Abs(Index = i)
    Next i
    SelIndex = Index + hS.Value
    
    With NewSelImage
        .FileName = picImage(Index).ToolTipText
        .IsResource = True
        Set .Picture = picImage(Index)
    End With
    
    cmdUpdate.Enabled = picImage(0).Visible
    cmdSave.Enabled = cmdUpdate.Enabled
    
End Sub

Private Sub GetImagesFromResource()

    Dim i As Integer
    Dim ff As Integer
    Dim FileName As String
    
    If Not IsReady Then Exit Sub

    On Error Resume Next
    
    ImageInFocus = False
    picContainer.Cls
    
    For i = 1 To UBound(Images)
        If i >= hS.Value Then
            If i - hS.Value >= MaxImages Then Exit For
            FileName = StatesPath + GetFileName(Split(Images(i), "::")(0))
            If (Not NoRender) And (Not ((Not SelImage.SupportsFlash) And IsFlash(FileName))) Then
                ff = FreeFile
                'Open FileName For Binary As #ff
                    'imgCode = Mid(Images(i), InStr(Images(i), "::") + 2)
                    'imgCode = Left(imgCode, Len(imgCode) - 2)
                    'SaveImageFile FileName, imgCode
                    SaveImageFile FileName, GetImgCode_FromRes(Split(Images(i), "::")(0))
                    'Put #ff, , imgCode
                'Close #ff
                Err.Clear
                If IsFlash(FileName) Then
                    Set picImage(i - hS.Value).Picture = LoadResPicture(200, vbResIcon)
                Else
                    If IsANI(FileName) Then
                        Set picImage(i - hS.Value).Picture = LoadResPicture(201, vbResIcon)
                    Else
                        Set picImage(i - hS.Value).Picture = OLLoadPicture(FileName)
                    End If
                End If
                If Err.number > 0 Then
                    Set picImage(i - hS.Value).Picture = LoadPicture()
                    picImage(i - hS.Value).BackColor = vbRed
                Else
                    picImage(i - hS.Value).BackColor = &HE0E0E0
                End If
                picImage(i - hS.Value).ToolTipText = Split(Images(i), "::")(0)
                lblImageName(i - hS.Value).caption = FitText(GetFileName(Split(Images(i), "::")(0)))
                picImage(i - hS.Value).BorderStyle = Abs(i = SelIndex)
    
                picImage(i - hS.Value).Visible = True
                With lblImageName(i - hS.Value)
                    .Left = picImage(i - hS.Value).Left + picImage(i - hS.Value).Width / 2 - .Width / 2
                    .Visible = True
                    .ForeColor = IIf(FileExists(picImage(i - hS.Value).ToolTipText), vbBlack, vbRed)
                End With
            End If
            
            If GetFileName(FileName) = GetFileName(SelImage.FileName) Then
                DoBevel picContainer, picImage(i - hS.Value), 2, RGB(20, 20, 20), RGB(20, 20, 20), RGB(245, 245, 245), RGB(245, 245, 245)
                If SelIndex = -1 Then SelIndex = i
                ImageInFocus = True
            End If
        End If
    Next i
    
    If Not NoRender Then
        For i = i - hS.Value To UBound(Images)
            picImage(i).Visible = False
            lblImageName(i).Visible = False
        Next i
    End If
    
End Sub

Private Function FitText(sStr As String) As String

    Dim i As Integer
    Dim nStr As String
    
    nStr = sStr
    
    For i = Len(nStr) To 1 Step -1
        lblS.caption = nStr
        If lblS.Width < 930 Then Exit For
        nStr = Left(nStr, i - 1)
    Next i
    
    FitText = nStr + IIf(nStr <> sStr, "…", vbNullString)

End Function

Private Sub picImage_DblClick(Index As Integer)

    cmdOK_Click

End Sub

Private Sub tmrInit_Timer()

    tmrInit.Enabled = False
    
    If Preferences.EnableUndoRedo Then
        InitDialog
    Else
        cmdFromFile_Click
    End If
    
    Form_Resize

End Sub

Private Sub LocalizeUI()

    caption = GetLocalizedStr(532)
    
    cmdFromFile.caption = GetLocalizedStr(533) + "   "
    cmdFromFile.Width = SetCtrlWidth(cmdFromFile) + 80
    
    cmdSave.caption = GetLocalizedStr(130) + "...   "
    cmdSave.Left = cmdFromFile.Left + cmdFromFile.Width + 90
    cmdSave.Width = SetCtrlWidth(cmdSave) + 80
    
    cmdUpdate.caption = GetLocalizedStr(673) + "   "
    cmdUpdate.Left = cmdSave.Left + cmdSave.Width + 90
    cmdUpdate.Width = SetCtrlWidth(cmdUpdate) + 80
    
    cmdOK.caption = GetLocalizedStr(186)
    cmdCancel.caption = GetLocalizedStr(187)
    
    If Preferences.language <> "eng" Then
        cmdOK.Width = SetCtrlWidth(cmdOK)
        cmdCancel.Width = SetCtrlWidth(cmdCancel)
    End If

End Sub
