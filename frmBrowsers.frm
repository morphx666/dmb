VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{9A2947B4-87EE-40EE-A3EF-32BDC32D5726}#1.0#0"; "xfxline3d.ocx"
Begin VB.Form frmBrowsers 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Available Browsers"
   ClientHeight    =   2700
   ClientLeft      =   5985
   ClientTop       =   5790
   ClientWidth     =   5970
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   Begin xfxLine3D.ucLine3D uc3DLine1 
      Height          =   30
      Left            =   30
      TabIndex        =   6
      Top             =   2175
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   53
   End
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   3735
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2025
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.CommandButton cmdDefault 
      Caption         =   "Set as &Default"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4620
      TabIndex        =   2
      Top             =   565
      Width           =   1290
   End
   Begin MSComctlLib.ImageList ilIcons 
      Left            =   2205
      Top             =   2025
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowsers.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4620
      TabIndex        =   7
      Top             =   2280
      Width           =   1290
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "De&lete"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4620
      TabIndex        =   4
      Top             =   1515
      Width           =   1290
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit..."
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4620
      TabIndex        =   3
      Top             =   1040
      Width           =   1290
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4620
      TabIndex        =   1
      Top             =   90
      Width           =   1290
   End
   Begin MSComctlLib.ListView lvBrowsers 
      Height          =   1770
      Left            =   75
      TabIndex        =   0
      Top             =   90
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   3122
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ilIcons"
      SmallIcons      =   "ilIcons"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "chName"
         Text            =   "Name"
         Object.Width           =   7285
      EndProperty
   End
End
Attribute VB_Name = "frmBrowsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()

    frmAddBrowser.Show vbModal
    GetBrowsers

End Sub

Private Sub cmdDefault_Click()

    Dim i As Integer
    
    For i = 1 To lvBrowsers.ListItems.count
        lvBrowsers.ListItems(i).Bold = (lvBrowsers.SelectedItem = lvBrowsers.ListItems(i))
        If lvBrowsers.ListItems(i).Bold Then SaveSetting App.EXEName, "Browsers", "Default", i
    Next i
    
    lvBrowsers.SetFocus

End Sub

Private Sub cmdDelete_Click()

    Dim i As Integer
    
    i = lvBrowsers.SelectedItem.Index - 1
    If i < (lvBrowsers.ListItems.count - 1) Then
        Do
            SaveSetting App.EXEName, "Browsers", "Name" & i, GetSetting(App.EXEName, "Browsers", "Name" & i + 1)
            SaveSetting App.EXEName, "Browsers", "Command" & i, GetSetting(App.EXEName, "Browsers", "Command" & i + 1)
            i = i + 1
        Loop Until LenB(GetSetting(App.EXEName, "Browsers", "Name" & i + 1, vbNullString)) = 0
    End If
    
    DeleteSetting App.EXEName, "Browsers", "Name" & i
    DeleteSetting App.EXEName, "Browsers", "Command" & i
    
    GetBrowsers

End Sub

Private Sub cmdEdit_Click()

    With frmAddBrowser
        .Caption = GetLocalizedStr(425)
        .txtName.Text = Left(lvBrowsers.SelectedItem.Text, Len(lvBrowsers.SelectedItem.Text) - Len(GetFileVersion(lvBrowsers.SelectedItem.Tag, True)) - 1)
        .txtFileName.Text = lvBrowsers.SelectedItem.Tag
        .cmdOK.Caption = GetLocalizedStr(426)
        .Show vbModal
    End With
    
    GetBrowsers

End Sub

Private Sub cmdOK_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    CenterForm Me
    SetupCharset Me
    LocalizeUI
    
    GetBrowsers

End Sub

Private Sub GetBrowsers()

    Dim i As Integer
    Dim bItm As ListItem
    Dim SelIdx As Integer
    
    If Not lvBrowsers.SelectedItem Is Nothing Then
        SelIdx = lvBrowsers.SelectedItem.Index - 1
    End If
    
    lvBrowsers.ListItems.Clear
    
    Set bItm = lvBrowsers.ListItems.Add(, , "Microsoft Internet Explorer (Internal)", , 1)
    
    i = 1
    Do Until LenB(GetSetting(App.EXEName, "Browsers", "Name" & i, vbNullString)) = 0
        Set bItm = lvBrowsers.ListItems.Add(, , GetSetting(App.EXEName, "Browsers", "Name" & i))
        bItm.Tag = GetSetting(App.EXEName, "Browsers", "Command" & i)
        
        If FileExists(bItm.Tag) Then
            bItm.Text = bItm.Text + " " + GetFileVersion(bItm.Tag, True)
            
            GetIcon picIcon, bItm.Tag
            ilIcons.ListImages.Add , , picIcon.Image
            bItm.SmallIcon = ilIcons.ListImages.count
        Else
            bItm.ForeColor = vbRed
        End If
        
        If SelIdx = i Then
            bItm.Selected = True
            bItm.EnsureVisible
        End If
        
        i = i + 1
    Loop
    
    If Val(GetSetting(App.EXEName, "Browsers", "Default", 1)) > lvBrowsers.ListItems.count Then
        SaveSetting App.EXEName, "Browsers", "Default", 1
    End If
    
    With lvBrowsers.ListItems(Val(GetSetting(App.EXEName, "Browsers", "Default", 1)))
        .Bold = True
        .Selected = True
        .EnsureVisible
    End With
    
    lvBrowsers_ItemClick lvBrowsers.SelectedItem
    
End Sub

Private Sub lvBrowsers_DblClick()

    cmdDefault_Click

End Sub

Private Sub lvBrowsers_ItemClick(ByVal Item As MSComctlLib.ListItem)

    cmdEdit.Enabled = Item.Index <> 1
    cmdDelete.Enabled = Item.Index <> 1

End Sub

Private Sub LocalizeUI()

    Caption = GetLocalizedStr(419)
    
    cmdAdd.Caption = GetLocalizedStr(420)
    cmdDefault.Caption = GetLocalizedStr(421)
    cmdEdit.Caption = GetLocalizedStr(422)
    cmdDelete.Caption = GetLocalizedStr(423)
    
    cmdOK.Caption = GetLocalizedStr(424)
    
End Sub
