VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{E1C6DB9D-BD4A-4A61-A759-0CED75D034BF}#43.0#0"; "SmartButton.ocx"
Begin VB.Form frmSound 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sound"
   ClientHeight    =   3705
   ClientLeft      =   6735
   ClientTop       =   4725
   ClientWidth     =   4095
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
   Icon            =   "frmSound.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog cDlg 
      Left            =   1035
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3075
      TabIndex        =   3
      Top             =   3210
      Width           =   900
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1935
      TabIndex        =   2
      Top             =   3210
      Width           =   900
   End
   Begin VB.Frame Frame1 
      Caption         =   "On Click"
      Height          =   1410
      Left            =   120
      TabIndex        =   6
      Top             =   1635
      Width           =   3855
      Begin VB.TextBox txtSound 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   165
         TabIndex        =   1
         Top             =   630
         Width           =   2475
      End
      Begin SmartButtonProject.SmartButton cmdBrowseSound 
         Height          =   360
         Index           =   1
         Left            =   2700
         TabIndex        =   12
         Top             =   585
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   635
         Picture         =   "frmSound.frx":014A
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
      Begin SmartButtonProject.SmartButton cmdPlay 
         Height          =   360
         Index           =   1
         Left            =   3165
         TabIndex        =   13
         Top             =   585
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   635
         Picture         =   "frmSound.frx":02A4
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
      Begin SmartButtonProject.SmartButton cmdRemoveSound 
         Height          =   240
         Index           =   1
         Left            =   1635
         TabIndex        =   8
         Top             =   975
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   423
         Caption         =   "Remove"
         Picture         =   "frmSound.frx":03FE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionLayout   =   4
         PictureLayout   =   3
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Audio File"
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
         Left            =   165
         TabIndex        =   7
         Top             =   420
         Width           =   690
      End
   End
   Begin VB.Frame frameOver 
      Caption         =   "Mouse Over"
      Height          =   1410
      Left            =   120
      TabIndex        =   4
      Top             =   105
      Width           =   3855
      Begin VB.TextBox txtSound 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   165
         TabIndex        =   0
         Top             =   630
         Width           =   2475
      End
      Begin SmartButtonProject.SmartButton cmdBrowseSound 
         Height          =   360
         Index           =   0
         Left            =   2715
         TabIndex        =   9
         Top             =   592
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   635
         Picture         =   "frmSound.frx":0798
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
      Begin SmartButtonProject.SmartButton cmdPlay 
         Height          =   360
         Index           =   0
         Left            =   3180
         TabIndex        =   10
         Top             =   592
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   635
         Picture         =   "frmSound.frx":08F2
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
      Begin SmartButtonProject.SmartButton cmdRemoveSound 
         Height          =   240
         Index           =   0
         Left            =   1635
         TabIndex        =   11
         Top             =   975
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   423
         Caption         =   "Remove"
         Picture         =   "frmSound.frx":0A4C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionLayout   =   4
         PictureLayout   =   3
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Audio File"
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
         Left            =   165
         TabIndex        =   5
         Top             =   420
         Width           =   690
      End
   End
End
Attribute VB_Name = "frmSound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBrowseSound_Click(Index As Integer)

    On Error GoTo ExitSub
    
    With cDlg
        .DialogTitle = "Select Sound File"
        .InitDir = Project.UserConfigs(Project.DefaultConfig).RootWeb
        .filter = SupportedAudioFiles
        .CancelError = True
        .Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly + cdlOFNLongNames + cdlOFNPathMustExist
        .ShowOpen
        txtSound(Index).Text = .FileName
    End With
    
ExitSub:
    Exit Sub

End Sub

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdOK_Click()

    If IsCommand(frmMain.tvMenus.SelectedItem.key) Then
        With MenuCmds(GetID)
            .Sound.OnMouseOver = txtSound(0).Text
            .Sound.OnClick = txtSound(1).Text
            frmMain.SaveState "Change " + .Name + " Sound"
        End With
    Else
        With MenuGrps(GetID)
            .Sound.OnMouseOver = txtSound(0).Text
            .Sound.OnClick = txtSound(1).Text
            frmMain.SaveState "Change " + .Name + " Sound"
        End With
    End If
    
    Unload Me

End Sub

Private Sub cmdPlay_Click(Index As Integer)

    PlaySound ByVal txtSound(Index).Text, 0&, SND_FILENAME Or SND_ASYNC Or SND_NOWAIT

End Sub

Private Sub cmdRemoveSound_Click(Index As Integer)

    txtSound(Index).Text = ""

End Sub

Private Sub Form_Load()

    CenterForm Me
    
    If IsCommand(frmMain.tvMenus.SelectedItem.key) Then
        With MenuCmds(GetID)
            txtSound(0).Text = .Sound.OnMouseOver
            txtSound(1).Text = .Sound.OnClick
            
            Caption = "Sounds - [" + MenuGrps(.Parent).Name + "/" + .Name + "]"
        End With
    Else
        With MenuGrps(GetID)
            txtSound(0).Text = .Sound.OnMouseOver
            txtSound(1).Text = .Sound.OnClick
            
            Caption = "Sounds - [" + .Name + "]"
        End With
    End If

End Sub

