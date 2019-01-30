VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPosOffsetAdvancedAdd 
   Caption         =   "Custom Offset Definition"
   ClientHeight    =   6975
   ClientLeft      =   4455
   ClientTop       =   4740
   ClientWidth     =   9795
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPosOffsetAdvancedAdd.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   9795
   Begin VB.Timer tmrInit 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   7485
      Top             =   975
   End
   Begin VB.Frame frameConAdv 
      Caption         =   "Condition"
      Height          =   3390
      Left            =   2700
      TabIndex        =   30
      Top             =   2100
      Width           =   6495
      Begin VB.Frame frameAdvCode 
         BorderStyle     =   0  'None
         Height          =   2625
         Left            =   195
         TabIndex        =   32
         Top             =   645
         Width           =   6165
         Begin VB.TextBox txtAdv 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2100
            Left            =   0
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   34
            Top             =   240
            Width           =   6090
         End
         Begin VB.Label lblAdvInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Type %f where you want DHTML Menu Builder to insert the offset values"
            ForeColor       =   &H80000011&
            Height          =   195
            Left            =   30
            TabIndex        =   35
            Top             =   2385
            Width           =   5265
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Type the condition using the JavaScript language"
            Height          =   195
            Left            =   0
            TabIndex        =   33
            Top             =   0
            Width           =   3540
         End
      End
      Begin MSComctlLib.TabStrip tsAvdOP 
         Height          =   3045
         Left            =   90
         TabIndex        =   31
         Top             =   270
         Width           =   6315
         _ExtentX        =   11139
         _ExtentY        =   5371
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   3
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Toolbars"
               Key             =   "tsToolbars"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Root Level Menus"
               Key             =   "tsRLM"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "SubMenus"
               Key             =   "tsSM"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame frameConAssist 
      Caption         =   "Condition"
      Height          =   3390
      Left            =   315
      TabIndex        =   3
      Top             =   1230
      Width           =   6495
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   195
         TabIndex        =   23
         Top             =   2190
         Width           =   6030
         Begin VB.OptionButton opNS 
            Caption         =   "Ignore Netscape 4"
            Height          =   195
            Index           =   2
            Left            =   3660
            TabIndex        =   26
            Top             =   0
            Value           =   -1  'True
            Width           =   1890
         End
         Begin VB.OptionButton opNS 
            Caption         =   "Not Netscape 4"
            Height          =   195
            Index           =   1
            Left            =   1680
            TabIndex        =   25
            Top             =   0
            Width           =   1425
         End
         Begin VB.OptionButton opNS 
            Caption         =   "Netscape 4"
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   24
            Top             =   0
            Width           =   1140
         End
      End
      Begin VB.Frame frameIE 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   195
         TabIndex        =   7
         Top             =   1215
         Width           =   6030
         Begin VB.OptionButton opIE 
            Caption         =   "Internet Explorer"
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   8
            Top             =   45
            Width           =   1545
         End
         Begin VB.OptionButton opIE 
            Caption         =   "Not Internet Explorer"
            Height          =   195
            Index           =   1
            Left            =   1680
            TabIndex        =   9
            Top             =   45
            Width           =   1845
         End
         Begin VB.OptionButton opIE 
            Caption         =   "Ignore Internet Explorer"
            Height          =   195
            Index           =   2
            Left            =   3660
            TabIndex        =   10
            Top             =   45
            Value           =   -1  'True
            Width           =   2070
         End
      End
      Begin VB.ComboBox cmbOS 
         Height          =   315
         ItemData        =   "frmPosOffsetAdvancedAdd.frx":058A
         Left            =   180
         List            =   "frmPosOffsetAdvancedAdd.frx":059A
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   525
         Width           =   1950
      End
      Begin VB.ComboBox cmbVerComp 
         Height          =   315
         ItemData        =   "frmPosOffsetAdvancedAdd.frx":05B8
         Left            =   180
         List            =   "frmPosOffsetAdvancedAdd.frx":05D1
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   2820
         Width           =   1950
      End
      Begin VB.TextBox txtVer 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2265
         TabIndex        =   29
         Text            =   "0"
         Top             =   2835
         Width           =   645
      End
      Begin VB.Frame frameSM 
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   195
         TabIndex        =   11
         Top             =   1500
         Width           =   6030
         Begin VB.OptionButton opSM 
            Caption         =   "Mozilla-based"
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   12
            Top             =   0
            Width           =   1260
         End
         Begin VB.OptionButton opSM 
            Caption         =   "Not Mozilla-based"
            Height          =   195
            Index           =   1
            Left            =   1680
            TabIndex        =   13
            Top             =   0
            Width           =   1560
         End
         Begin VB.OptionButton opSM 
            Caption         =   "Ignore Mozilla-based"
            Height          =   195
            Index           =   2
            Left            =   3660
            TabIndex        =   14
            Top             =   0
            Value           =   -1  'True
            Width           =   1785
         End
      End
      Begin VB.Frame frameOP 
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   195
         TabIndex        =   15
         Top             =   1740
         Width           =   6030
         Begin VB.OptionButton opOP 
            Caption         =   "Opera"
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   16
            Top             =   0
            Width           =   750
         End
         Begin VB.OptionButton opOP 
            Caption         =   "Not Opera"
            Height          =   195
            Index           =   1
            Left            =   1680
            TabIndex        =   17
            Top             =   0
            Width           =   1050
         End
         Begin VB.OptionButton opOP 
            Caption         =   "Ignore Opera"
            Height          =   195
            Index           =   2
            Left            =   3660
            TabIndex        =   18
            Top             =   0
            Value           =   -1  'True
            Width           =   1275
         End
      End
      Begin VB.Frame frameKQ 
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   195
         TabIndex        =   19
         Top             =   1965
         Width           =   6030
         Begin VB.OptionButton opKQ 
            Caption         =   "Safari"
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   20
            Top             =   0
            Width           =   720
         End
         Begin VB.OptionButton opKQ 
            Caption         =   "Not Safari"
            Height          =   195
            Index           =   1
            Left            =   1680
            TabIndex        =   21
            Top             =   0
            Width           =   1020
         End
         Begin VB.OptionButton opKQ 
            Caption         =   "Ignore Safari"
            Height          =   195
            Index           =   2
            Left            =   3660
            TabIndex        =   22
            Top             =   0
            Value           =   -1  'True
            Width           =   1245
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "If the operating system is"
         Height          =   195
         Left            =   180
         TabIndex        =   4
         Top             =   285
         Width           =   1845
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "And the browser is"
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   1035
         Width           =   1350
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "And the browser version is"
         Height          =   195
         Left            =   180
         TabIndex        =   27
         Top             =   2595
         Width           =   1920
      End
   End
   Begin VB.Frame frameOffsets 
      Caption         =   "Offsets"
      Height          =   1425
      Left            =   180
      TabIndex        =   36
      Top             =   4950
      Width           =   6780
      Begin VB.TextBox txtTV 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   165
         TabIndex        =   46
         Text            =   "0"
         Top             =   900
         Width           =   480
      End
      Begin VB.TextBox txtTH 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   165
         TabIndex        =   40
         Text            =   "0"
         Top             =   525
         Width           =   480
      End
      Begin VB.TextBox txtSV 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4725
         TabIndex        =   50
         Text            =   "0"
         Top             =   900
         Width           =   480
      End
      Begin VB.TextBox txtSH 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4725
         TabIndex        =   45
         Text            =   "0"
         Top             =   525
         Width           =   480
      End
      Begin VB.TextBox txtRH 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2310
         TabIndex        =   44
         Text            =   "0"
         Top             =   525
         Width           =   480
      End
      Begin VB.TextBox txtRV 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2310
         TabIndex        =   48
         Text            =   "0"
         Top             =   900
         Width           =   480
      End
      Begin VB.Label label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pixels Vertically"
         Height          =   195
         Left            =   720
         TabIndex        =   47
         Top             =   945
         Width           =   1095
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pixels Horizontally"
         Height          =   195
         Left            =   720
         TabIndex        =   43
         Top             =   585
         Width           =   1290
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Toolbars"
         Height          =   195
         Left            =   165
         TabIndex        =   37
         Top             =   285
         Width           =   615
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SubMenus"
         Height          =   195
         Left            =   4725
         TabIndex        =   39
         Top             =   285
         Width           =   735
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Root Level Menus"
         Height          =   195
         Left            =   2310
         TabIndex        =   38
         Top             =   285
         Width           =   1275
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pixels Vertically"
         Height          =   195
         Left            =   5280
         TabIndex        =   51
         Top             =   945
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pixels Horizontally"
         Height          =   195
         Left            =   5280
         TabIndex        =   42
         Top             =   570
         Width           =   1290
      End
      Begin VB.Label lblRootH 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pixels Horizontally"
         Height          =   195
         Left            =   2865
         TabIndex        =   41
         Top             =   570
         Width           =   1290
      End
      Begin VB.Label lblRootV 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pixels Vertically"
         Height          =   195
         Left            =   2865
         TabIndex        =   49
         Top             =   945
         Width           =   1095
      End
   End
   Begin MSComctlLib.TabStrip tsMode 
      Height          =   4020
      Left            =   180
      TabIndex        =   2
      Top             =   765
      Width           =   6780
      _ExtentX        =   11959
      _ExtentY        =   7091
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Assistant"
            Key             =   "tsAssistant"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Advanced"
            Key             =   "tsAdvanced"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtDescription 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   345
      Width           =   3795
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6060
      TabIndex        =   53
      Top             =   6495
      Width           =   900
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4965
      TabIndex        =   52
      Top             =   6495
      Width           =   900
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   795
   End
End
Attribute VB_Name = "frmPosOffsetAdvancedAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ADVt As String
Dim ADVr As String
Dim ADVs As String

Private WithEvents msgSubClass As xfxSC
Attribute msgSubClass.VB_VarHelpID = -1
Private IsResizing As Boolean

Private Sub cmbVerComp_Click()

    txtVer.Enabled = (cmbVerComp.ListIndex > 0)

End Sub

Private Sub cmdCancel_Click()

    curOffsetStr = ""
    Unload Me

End Sub

Private Sub cmdOK_Click()

    Dim sStr As String
    
    sStr = txtDescription.Text + "@"
    
     Select Case tsMode.SelectedItem.key
        Case "tsAssistant"
            sStr = sStr + "W@"
            sStr = sStr + "OS:" & cmbOS.ListIndex & "@"
    
            sStr = sStr + getBrowserVal(opIE, "IE")
            sStr = sStr + getBrowserVal(opSM, "SM")
            sStr = sStr + getBrowserVal(opOP, "OP")
            sStr = sStr + getBrowserVal(opKQ, "KQ")
            sStr = sStr + getBrowserVal(opNS, "NS")
            
            sStr = sStr + "VER:" & cmbVerComp.ListIndex & ":" & txtVer.Text + "@"
        Case "tsAdvanced"
            sStr = sStr + "A@"
            sStr = sStr + "T:" + Replace(Replace(ADVt, vbCrLf, "\n"), ":", Chr(255)) + "@"
            sStr = sStr + "R:" + Replace(Replace(ADVr, vbCrLf, "\n"), ":", Chr(255)) + "@"
            sStr = sStr + "S:" + Replace(Replace(ADVs, vbCrLf, "\n"), ":", Chr(255)) + "@"
    End Select
    
    sStr = sStr + "TH:" + txtTH.Text + "@"
    sStr = sStr + "TV:" + txtTV.Text + "@"
    
    sStr = sStr + "RH:" + txtRH.Text + "@"
    sStr = sStr + "RV:" + txtRV.Text + "@"
    
    sStr = sStr + "SH:" + txtSH.Text + "@"
    sStr = sStr + "SV:" + txtSV.Text
    
    curOffsetStr = sStr
    
    Unload Me

End Sub

Private Sub LoadCurOffsetParams()

    Dim o() As String

    o = Split(curOffsetStr, "@")
    txtDescription.Text = o(0)
    Select Case o(1)
        Case "W"
            cmbOS.ListIndex = Split(o(2), ":")(1)
            
            opIE(Split(o(3), ":")(1)).Value = True
            opSM(Split(o(4), ":")(1)).Value = True
            opOP(Split(o(5), ":")(1)).Value = True
            opKQ(Split(o(6), ":")(1)).Value = True
            opNS(Split(o(7), ":")(1)).Value = True
            
            cmbVerComp.ListIndex = Split(o(8), ":")(1)
            txtVer.Text = Split(o(8), ":")(2)
            
            txtTH.Text = Split(o(9), ":")(1)
            txtTV.Text = Split(o(10), ":")(1)
            
            txtRH.Text = Split(o(11), ":")(1)
            txtRV.Text = Split(o(12), ":")(1)
            
            txtSH.Text = Split(o(13), ":")(1)
            txtSV.Text = Split(o(14), ":")(1)
        Case "A"
            ADVt = Replace(Replace(Split(o(2), ":")(1), "\n", vbCrLf), Chr(255), ":")
            ADVr = Replace(Replace(Split(o(3), ":")(1), "\n", vbCrLf), Chr(255), ":")
            ADVs = Replace(Replace(Split(o(4), ":")(1), "\n", vbCrLf), Chr(255), ":")
            
            txtTH.Text = Split(o(5), ":")(1)
            txtTV.Text = Split(o(6), ":")(1)
            
            txtRH.Text = Split(o(7), ":")(1)
            txtRV.Text = Split(o(8), ":")(1)
            
            txtSH.Text = Split(o(9), ":")(1)
            txtSV.Text = Split(o(10), ":")(1)
            
            tsMode.Tabs("tsAdvanced").Selected = True
            tsMode_Click
            
            If LenB(ADVt) <> 0 Then
                tsAvdOP.Tabs("tsToolbars").Selected = True
            Else
                If LenB(ADVr) <> 0 Then
                    tsAvdOP.Tabs("tsRLM").Selected = True
                Else
                    If LenB(ADVs) <> 0 Then tsAvdOP.Tabs("tsSM").Selected = True
                End If
            End If
            tsAvdOP_Click
        End Select

End Sub

Private Function getBrowserVal(op As Variant, sStr As String) As String

    Dim c As String

    c = sStr + ":"
    If op(0).Value Then c = c + "0"
    If op(1).Value Then c = c + "1"
    If op(2).Value Then c = c + "2"
    
    getBrowserVal = c + "@"

End Function

Private Sub Form_Load()

    txtAdv.Enabled = False
    
    If Val(GetSetting(App.EXEName, "PosOffAdvPos", "X")) = 0 Then
        Width = 7215
        Height = 7410
        
        CenterForm Me
    Else
        Left = GetSetting(App.EXEName, "PosOffAdvPos", "X")
        Top = GetSetting(App.EXEName, "PosOffAdvPos", "Y")
        Width = GetSetting(App.EXEName, "PosOffAdvPos", "W")
        Height = GetSetting(App.EXEName, "PosOffAdvPos", "H")
    End If
    
    ADVt = ""
    ADVr = ""
    ADVs = ""

    cmbOS.ListIndex = 0
    cmbVerComp.ListIndex = 0
    
    With frameConAssist
        frameConAdv.Move .Left, .Top
        .ZOrder 0
    End With
    FixCtrls4Skin Me
    
    tmrInit.Enabled = True
    
    If Not IsDebug Then SetupSubclassing True

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    SaveSetting App.EXEName, "PosOffAdvPos", "X", Left
    SaveSetting App.EXEName, "PosOffAdvPos", "Y", Top
    SaveSetting App.EXEName, "PosOffAdvPos", "W", Width
    SaveSetting App.EXEName, "PosOffAdvPos", "H", Height

    SetupSubclassing False

End Sub

Private Sub Form_Resize()

    Dim tppx As Integer
    Dim tppy As Integer
    Dim leftBorder As Integer
    Dim topBorder As Integer
    
    IsResizing = True
    
    If Me.Width < 7185 Then Me.Width = 7185
    If Me.Height < 7410 Then Me.Width = 7410
    
    tppx = Screen.TwipsPerPixelX
    tppy = Screen.TwipsPerPixelY
    
    leftBorder = GetClientLeft(Me.hwnd)
    topBorder = GetClientTop(Me.hwnd)
    
    tsMode.Width = Me.Width - tsMode.Left * 2 - leftBorder
    frameConAssist.Width = Me.Width - frameConAssist.Left * 2 - leftBorder
    frameOffsets.Width = Me.Width - frameOffsets.Left * 2 - leftBorder
    frameOffsets.Top = Me.Height - frameOffsets.Height - topBorder - cmdCancel.Height - 15 * tppy
    tsMode.Height = frameOffsets.Top - tsMode.Top - 10 * tppy
    frameConAssist.Height = tsMode.Height - 42 * tppy
    
    frameConAdv.Width = frameConAssist.Width
    frameConAdv.Height = frameConAssist.Height
    
    tsAvdOP.Width = frameConAdv.Width - tsAvdOP.Left * 2
    tsAvdOP.Height = frameConAdv.Height - tsAvdOP.Top - 7 * tppy
    
    frameAdvCode.Width = tsAvdOP.Width - frameAdvCode.Left
    frameAdvCode.Height = tsAvdOP.Height - frameAdvCode.Top + 14 * tppy
    
    txtAdv.Width = frameAdvCode.Width - txtAdv.Left * 2
    txtAdv.Height = frameAdvCode.Height - txtAdv.Top * 2 - tppy
    lblAdvInfo.Top = txtAdv.Top + txtAdv.Height + 2 * tppy
    
    cmdCancel.Move frameOffsets.Left + frameOffsets.Width - cmdCancel.Width, Me.Height - topBorder - cmdCancel.Height - 7 * tppy
    cmdOK.Move cmdCancel.Left - cmdOK.Width - 10 * tppx, cmdCancel.Top
    
    IsResizing = False

End Sub

Private Sub tmrInit_Timer()

    tmrInit.Enabled = False
    If LenB(curOffsetStr) <> 0 Then LoadCurOffsetParams

End Sub

Private Sub tsAvdOP_Click()

    Select Case tsAvdOP.SelectedItem.key
        Case "tsToolbars": txtAdv.Text = ADVt
        Case "tsRLM": txtAdv.Text = ADVr
        Case "tsSM": txtAdv.Text = ADVs
    End Select
    
    txtAdv.SetFocus

End Sub

Private Sub tsMode_Click()

    Select Case tsMode.SelectedItem.key
        Case "tsAssistant"
            frameConAssist.ZOrder 0
            txtAdv.Enabled = False
        Case "tsAdvanced"
            txtAdv.Enabled = True
            frameConAdv.ZOrder 0
            tsAvdOP_Click
    End Select

End Sub

Private Sub txtAdv_Change()

    Select Case tsAvdOP.SelectedItem.key
        Case "tsToolbars": ADVt = txtAdv.Text
        Case "tsRLM": ADVr = txtAdv.Text
        Case "tsSM": ADVs = txtAdv.Text
    End Select

End Sub

Private Sub txtAdv_GotFocus()

    Dim c As Control
    
    On Error Resume Next
    
    For Each c In Me.Controls
        c.TabStop = False
    Next c

End Sub

Private Sub txtAdv_LostFocus()

    Dim c As Control
    
    On Error Resume Next
    
    For Each c In Me.Controls
        If (TypeOf c Is Button Or TypeOf c Is TextBox Or TypeOf c Is TextBox) And c.Enabled And c.Visible Then c.TabStop = True
    Next c

End Sub

Private Sub txtDescription_Change()

    cmdOK.Enabled = (LenB(txtDescription.Text) <> 0)

End Sub

Private Sub txtRH_GotFocus()

    SelAll txtRH

End Sub

Private Sub txtRH_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 45 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtRV_GotFocus()

    SelAll txtRV

End Sub

Private Sub txtRV_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 45 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtSH_GotFocus()

    SelAll txtSH

End Sub

Private Sub txtSH_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 45 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtSV_GotFocus()

    SelAll txtSV

End Sub

Private Sub txtSV_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 45 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtTH_GotFocus()

    SelAll txtTH

End Sub

Private Sub txtTH_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 45 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtTV_GotFocus()

    SelAll txtTV

End Sub

Private Sub txtTV_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 45 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtVer_GotFocus()

    SelAll txtVer

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF1 Then showHelp "dialogs/menus_offset_cusoff_add.htm"

End Sub

Private Sub SetupSubclassing(scState As Boolean)

    If msgSubClass Is Nothing Then Set msgSubClass = New xfxSC
    
    frmPOAHWND = Me.hwnd
    msgSubClass.SubClassHwnd Me.hwnd, scState

End Sub

Private Sub msgSubClass_NewMessage(ByVal hwnd As Long, uMsg As Long, wParam As Long, lParam As Long, Cancel As Boolean, UseRetVal As Boolean, RetVal As Long)

    If IsResizing Then Exit Sub

    Cancel = HandleSubclassMsg(hwnd, uMsg, wParam, lParam)

End Sub

