VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DMB's Objects Tester"
   ClientHeight    =   3690
   ClientLeft      =   6270
   ClientTop       =   4455
   ClientWidth     =   5460
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   5460
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy to Clipboard"
      Height          =   345
      Left            =   60
      TabIndex        =   4
      Top             =   3285
      Width           =   1635
   End
   Begin VB.TextBox txtLog 
      Height          =   2850
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   330
      Width           =   5340
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   345
      Left            =   3465
      TabIndex        =   1
      Top             =   3285
      Width           =   870
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   345
      Left            =   4545
      TabIndex        =   0
      Top             =   3285
      Width           =   870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Log Output"
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   75
      Width           =   810
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()

    Unload Me

End Sub

Private Sub cmdCopy_Click()

    Clipboard.Clear
    Clipboard.SetText txtLog.Text, vbCFText

End Sub

Private Sub cmdStart_Click()

    txtLog.Text = ""
    
    txtLog.Text = txtLog.Text + ln + "COM Components" + ln

    TestObject "Engine.CEngine"
    TestObject "SSubTimer6.CTimer"
    TestObject "SmartSubClassLib.SmartSubClass"
    TestObject "TipsSystem.CTips"
    TestObject "IconMenu6.cIconMenu"
    TestObject "MSComCtl2.DTPicker.2"
    TestObject "MSComCtl2.MonthView.2"
    TestObject "MSComCtl2.UpDown.2"
    TestObject "MSComCtl2.Animation.2"
    TestObject "MSComCtl2.FlatScrollBar.2"
    TestObject "href.uchref1"
    TestObject "xfxLine3D.ucLine3D"
    TestObject "ColorPicker.ucColorPicker"
    TestObject "IContainer.IconContainer"
    TestObject "DMBSampleControl.ucDMBSampleCtrl"
    'TestObject "SmartButtonProject.SmartButton"
    'TestObject "VBSmartXPMenu.SmartMenuXP"
    TestObject "xfxFormShaper.FormShaper"
    TestObject "MSComctlLib.ImageListCtrl.2"
    TestObject "MSComctlLib.ProgCtrl.2"
    TestObject "MSComctlLib.Toolbar.2"
    TestObject "MSComctlLib.SBarCtrl.2"
    TestObject "MSComctlLib.ListViewCtrl.2"
    TestObject "MSComctlLib.TreeCtrl.2"
    TestObject "MSComctlLib.ImageComboCtl.2"
    TestObject "MSComctlLib.Slider.2"
    TestObject "MSComDlg.CommonDialog.1"
    TestObject "InetCtls.Inet.1"
    
    txtLog.Text = txtLog.Text + ln + "DHTML Menu Builder Applications" + ln
    
    chkFile "dmb.exe"
    chkFile "dmbc.exe"
    chkFile "presetinstaller.exe"
    chkFile "addininstaller.exe"
    
    txtLog.Text = txtLog.Text + ln + "DHTML Menu Builder Libraries" + ln
    
    chkFile "engine.dll"
    chkFile "tsys.dll"
    chkFile "dmbsamplecontrol.ocx"
    chkFile "xfxslider.ocx"
    chkFile "xfxbinimg.dll"
    
    txtLog.Text = txtLog.Text + ln + "xFX JumpStart Shared Components" + ln
    
    chkFile "href.ocx"
    chkFile "line3d.ocx"
    chkFile "colorpicker.ocx"
    chkFile "icontainer.ocx"
    
    txtLog.Text = txtLog.Text + ln + "VBSmart Shared Components" + ln
    
    chkFile "smartbutton.ocx"
    chkFile "smartsubclass.dll"
    chkFile "smartmenuxp.dll"
    chkFile "smartmenuxp.ocx"
    
    txtLog.Text = txtLog.Text + ln + "Microsoft Shared Components" + ln
    
    chkFile "mscomct2.ocx"
    chkFile "mscomctl.ocx"
    chkFile "msinet.ocx"
    chkFile "msvbvm60.dll"
    
    txtLog.Text = txtLog.Text + ln + "Image Handling Shared Components" + ln
    
    chkFile "gflax170.dll"
    chkFile "libgfl170.dll"
    chkFile "libgfle170.dll"

End Sub

Private Function ln() As String

    ln = vbCrLf + String(txtLog.Width / (5 * 15), "-") + vbCrLf

End Function

Private Sub chkFile(FileName As String)

    Dim sysDir As String
    Dim dmbDir As String
    Dim fn As String
    Dim r As String
    
    sysDir = GetSystemDir + "\"
    dmbDir = GetSetting("DMB", "RegInfo", "InstallPath")
    If FileExists(sysDir + FileName) Then
        fn = sysDir + FileName
    Else
        If FileExists(dmbDir + FileName) Then
            fn = dmbDir + FileName
        End If
    End If
    
    If fn = "" Then
        r = FileName + ": Failed => File does not exist"
    Else
        r = FileName + ": Passed " + vbCrLf + _
            vbTab + "Version: " + GetFileVersion(fn, False) + vbCrLf + _
            vbTab + "Size: " + FormatNumber(FileLen(fn), 0, , , vbTrue)
    End If
    txtLog.Text = txtLog.Text + vbTab + r + ln

End Sub

Private Sub TestObject(className As String)

    On Error Resume Next
    
    Dim obj As Object
    Dim r As String
    
    If className = "" Then Exit Sub
    
    Err.Clear
    
    Set obj = CreateObject(className)
    
    If Err.Number > 0 Then
        r = className + ": Failed => " + "Error: (" & Err.Number & ") " + Err.Description
    Else
        r = className + ": Passed"
    End If
    txtLog.Text = txtLog.Text + vbTab + r + ln
    
    Set obj = Nothing

End Sub

Private Sub Form_Load()

    With Screen
        Move (.Width - Width) / 2, (.Height - Height) / 2
    End With

End Sub
