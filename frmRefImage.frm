VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E1C6DB9D-BD4A-4A61-A759-0CED75D034BF}#43.0#0"; "SmartButton.ocx"
Object = "{9A2947B4-87EE-40EE-A3EF-32BDC32D5726}#1.0#0"; "xfxline3d.ocx"
Begin VB.Form frmRefImage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Reference Image"
   ClientHeight    =   5790
   ClientLeft      =   4620
   ClientTop       =   4155
   ClientWidth     =   7470
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
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkEnableRename 
      Height          =   285
      Left            =   3045
      Picture         =   "frmRefImage.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4515
      Width           =   345
   End
   Begin VB.Timer tmrInit 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3105
      Top             =   5400
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   5400
      TabIndex        =   9
      Top             =   5250
      Width           =   900
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6450
      TabIndex        =   8
      Top             =   5250
      Width           =   900
   End
   Begin VB.TextBox txtImageName 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   4515
      Width           =   1785
   End
   Begin VB.PictureBox picImage 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   840
      ScaleHeight     =   495
      ScaleWidth      =   375
      TabIndex        =   5
      Top             =   5325
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSComctlLib.ListView lvImages 
      Height          =   3435
      Left            =   150
      TabIndex        =   4
      Top             =   1050
      Width           =   7200
      _ExtentX        =   12700
      _ExtentY        =   6059
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ilImages"
      SmallIcons      =   "ilImages"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "File Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Image Name"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList ilImages 
      Left            =   4740
      Top             =   165
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
   End
   Begin VB.TextBox txtDoc 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   150
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   3660
   End
   Begin SmartButtonProject.SmartButton cmdSelRefImage 
      Height          =   315
      Left            =   3885
      TabIndex        =   2
      Top             =   345
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   556
      Picture         =   "frmRefImage.frx":038A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin xfxLine3D.ucLine3D uc3DLine1 
      Height          =   30
      Left            =   75
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5085
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   53
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Image Name"
      Height          =   195
      Left            =   150
      TabIndex        =   6
      Top             =   4560
      Width           =   900
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Selec the Image to attach the '%t' toolbar"
      Height          =   195
      Left            =   150
      TabIndex        =   3
      Top             =   825
      Width           =   3030
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Document Containing the Reference Image"
      Height          =   195
      Left            =   150
      TabIndex        =   0
      Top             =   135
      Width           =   3105
   End
End
Attribute VB_Name = "frmRefImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkEnableRename_Click()

    With txtImageName
        If chkEnableRename.Value = vbChecked Then
        
            With TipsSys
                .CanDisable = False
                .TipTitle = "Warning about changing the name of an image"
                .Tip = "It is highly recommended that you use your preferred HTML editor to change the name parameter of your images." + vbCrLf + _
                    "Although this feature has been tested under many different scenarios there's the possibility that your HTML document could become corrupted." + vbCrLf + vbCrLf + _
                    "Always make backup copies of the documents you edit using this tool."
                .Show
            End With
        
            .BackColor = &H80000005
            .ForeColor = &H80000008
            .Locked = False
            .SetFocus
        Else
            .BackColor = &HE0E0E0
            .ForeColor = &H80000008
            .Locked = True
            lvImages.SetFocus
        End If
    End With

End Sub

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdOK_Click()

    Project.ToolBar.AttachTo = txtImageName.Text
    Unload Me

End Sub

Private Sub cmdSelRefImage_Click()

    On Error GoTo ExitSub
    
    With frmMain.cDlg
        .DialogTitle = "Select the document containing the reference image"
        .InitDir = GetRealLocal.RootWeb
        .filter = SupportedHTMLDocs
        .CancelError = True
        .Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly + cdlOFNLongNames + cdlOFNPathMustExist
        .ShowOpen
        txtDoc.Text = .FileName
    End With
    
ExitSub:
    Exit Sub

End Sub

Private Sub Form_Load()

    SetupCharset Me
    CenterForm Me

    lblInfo.Caption = Replace(lblInfo.Caption, "%t", Project.ToolBar.Name)
    tmrInit.Enabled = True

End Sub

Private Sub lvImages_ItemClick(ByVal Item As MSComctlLib.ListItem)

    txtImageName.Text = Item.SubItems(1)
    
End Sub

Private Sub tmrInit_Timer()

    tmrInit.Enabled = False
    txtDoc.Text = GetRealLocal.HotSpotEditor.HotSpotsFile

End Sub

Private Sub txtDoc_Change()

    LoadImages

End Sub

Private Sub LoadImage(imgSrc As String)

    If LenB(imgSrc) <> 0 Then
        imgSrc = GetFilePath(txtDoc.Text) + imgSrc
        imgSrc = SetSlashDir(imgSrc, sdBack)
        imgSrc = RemoveDoubleSlashes(imgSrc)
        On Error Resume Next
        picImage.Picture = LoadPicture(imgSrc)
    End If
    
End Sub

Private Sub LoadImages()
    
    Dim atags() As String
    Dim p As Integer
    Dim fCode As String
    Dim i As Integer
    
    On Error GoTo ExitSub
    
    ReDim atags(0)
    
    lvImages.ListItems.Clear
    txtImageName.Text = ""
    
    fCode = LoadFile(txtDoc.Text)
    
    fCode = Replace(fCode, "<IMG ", "<img ")
    fCode = Replace(fCode, "<img" + vbCrLf, "<img ")
    fCode = Replace(fCode, "<IMG" + vbCrLf, "<img ")
    atags = Split(fCode, "<img ")
    
    If UBound(atags) > 0 Then
        For i = 1 To UBound(atags)
            p = InStr(atags(i), ">")
            If p > 0 Then
                LoadImage GetParamVal(atags(i), "src")
                ilImages.ListImages.Add , , picImage.Image
                With lvImages.ListItems.Add(, , GetFileName(GetParamVal(atags(i), "src")), ilImages.ListImages.Count)
                    .SubItems(1) = GetParamVal(atags(i), "name")
                    .Tag = atags(i)
                    If .SubItems(1) <> "" And .SubItems(1) = Project.ToolBar.AttachTo Then
                        .Bold = True
                        .Selected = True
                        txtImageName.Text = .SubItems(1)
                    End If
                    .SmallIcon = .Icon
                End With
            End If
        Next i
    End If
    
    lvImages.ColumnHeaders(1).Width = lvImages.Width / 2
    lvImages.ColumnHeaders(2).Width = lvImages.Width / 2 - 28 * 15
    lvImages.SetFocus
    
    If Not lvImages.SelectedItem Is Nothing Then lvImages.SelectedItem.EnsureVisible
    
ExitSub:

End Sub

Private Sub txtImageName_Change()

    Dim atags() As String
    Dim fCode As String
    Dim i As Integer
    
    On Error GoTo ExitSub
    
    ReDim atags(0)
    
    If lvImages.SelectedItem Is Nothing Then
        txtImageName.Text = ""
        GoTo ExitSub
    End If
    
    If txtImageName.Text <> "" And chkEnableRename.Value = vbChecked Then
        fCode = LoadFile(txtDoc.Text)
        
        fCode = Replace(fCode, "<IMG ", "<img ")
        fCode = Replace(fCode, "<img" + vbCrLf, "<img ")
        fCode = Replace(fCode, "<IMG" + vbCrLf, "<img ")
        atags = Split(fCode, "<img ")
        
        If UBound(atags) > 0 Then
            For i = 1 To UBound(atags)
                If atags(i) = lvImages.SelectedItem.Tag Then
                    If GetParamVal(atags(i), "name") = txtImageName.Text Then GoTo ExitSub
                    atags(i) = ChangeParamVal(atags(i), "name", txtImageName.Text, True)
                    lvImages.SelectedItem.Tag = atags(i)
                    Exit For
                End If
            Next i
        End If
        
        SaveFile txtDoc.Text, Join(atags, "<img ")
    End If
    
ExitSub:
    If Not lvImages.SelectedItem Is Nothing Then lvImages.SelectedItem.SubItems(1) = txtImageName.Text
    cmdOK.Enabled = (txtImageName.Text <> "")

End Sub
