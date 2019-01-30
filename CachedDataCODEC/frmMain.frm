VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cached Data CODEC"
   ClientHeight    =   5370
   ClientLeft      =   4965
   ClientTop       =   5550
   ClientWidth     =   6180
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
   ScaleHeight     =   5370
   ScaleWidth      =   6180
   Begin VB.CheckBox chkInflate 
      Caption         =   "Infalte"
      Height          =   195
      Left            =   150
      TabIndex        =   16
      Top             =   3720
      Width           =   990
   End
   Begin VB.CommandButton cemFromCache 
      Caption         =   "From Cache"
      Height          =   525
      Left            =   150
      TabIndex        =   15
      Top             =   2880
      Width           =   990
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Report"
      Height          =   420
      Left            =   1305
      TabIndex        =   14
      Top             =   4785
      Width           =   1305
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   420
      Left            =   4620
      TabIndex        =   13
      Top             =   4785
      Width           =   1305
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Report"
      Height          =   420
      Left            =   3180
      TabIndex        =   12
      Top             =   4785
      Width           =   1305
   End
   Begin VB.CommandButton cmdGetFromRegistry 
      Caption         =   "From Registry"
      Height          =   525
      Left            =   150
      TabIndex        =   11
      Top             =   2295
      Width           =   990
   End
   Begin VB.CommandButton cmdGetHDSerial 
      Caption         =   "Auto"
      Height          =   285
      Left            =   5430
      TabIndex        =   10
      Top             =   1365
      Width           =   495
   End
   Begin VB.TextBox txtCachedData 
      BeginProperty Font 
         Name            =   "Monotype.com"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2610
      Left            =   1305
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   2025
      Width           =   4620
   End
   Begin VB.TextBox txtSerial 
      BeginProperty Font 
         Name            =   "Monotype.com"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1305
      TabIndex        =   7
      Top             =   1365
      Width           =   4065
   End
   Begin VB.TextBox txtCompany 
      BeginProperty Font 
         Name            =   "Monotype.com"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1305
      TabIndex        =   5
      Top             =   990
      Width           =   4065
   End
   Begin VB.TextBox txtUserName 
      BeginProperty Font 
         Name            =   "Monotype.com"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1305
      TabIndex        =   3
      Top             =   615
      Width           =   4065
   End
   Begin VB.TextBox txtUnlockCode 
      BeginProperty Font 
         Name            =   "Monotype.com"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1305
      TabIndex        =   1
      Top             =   240
      Width           =   4065
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cached Data"
      Height          =   195
      Left            =   180
      TabIndex        =   8
      Top             =   2025
      Width           =   930
   End
   Begin VB.Line Line1 
      X1              =   165
      X2              =   6045
      Y1              =   1830
      Y2              =   1830
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Serial"
      Height          =   195
      Left            =   180
      TabIndex        =   6
      Top             =   1410
      Width           =   390
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Company"
      Height          =   195
      Left            =   180
      TabIndex        =   4
      Top             =   1035
      Width           =   675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      Height          =   195
      Left            =   180
      TabIndex        =   2
      Top             =   660
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unlock Code"
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   285
      Width           =   885
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private disableAutoEncode As Boolean

Private Sub cemFromCache_Click()

    FromCache

End Sub

Private Sub chkInflate_Click()

    cmdGetFromRegistry_Click

End Sub

Private Sub cmdClose_Click()

    Unload Me

End Sub

Private Sub cmdGetFromRegistry_Click()

    DecodeFromRegistry
    txtCachedData.SetFocus

End Sub

Private Sub DecodeFromRegistry()
  
    DoDecode GetSetting("DMB", "RegInfo", "CacheData")

End Sub

Private Sub cmdGetHDSerial_Click()

    txtSerial.Text = GetHDSerial
    DoEncode
    txtSerial.SetFocus

End Sub

Private Sub cmdLoad_Click()

    Dim r As String
    
    If FileExists(App.Path + "\cdcreport.dat") = False Then Exit Sub
        
    r = LoadFile(App.Path + "\cdcreport.dat")
    r = Inflate(HEX2Str(r))
    
    disableAutoEncode = True
    
    Dim l() As String
    l = Split(r, "|")
    
    txtUnlockCode.Text = l(0)
    txtUserName.Text = l(1)
    txtCompany.Text = l(2)
    txtSerial.Text = l(3)
    txtCachedData.Text = l(4)
    
    disableAutoEncode = False
    
    MsgBox "Report successfully loaded", vbInformation + vbOKOnly, "CachedDataCODEC Report"

End Sub

Private Sub cmdSave_Click()

    Dim r As String
    r = txtUnlockCode.Text + "|" + _
         txtUserName.Text + "|" + _
         txtCompany.Text + "|" + _
         txtSerial.Text + "|" + _
         txtCachedData.Text
        
    r = Str2HEX(Deflate(r))
    SaveFile App.Path + "\cdcreport.dat", r
    
    MsgBox "Report successfully saved", vbInformation + vbOKOnly, "CachedDataCODEC Report"

End Sub

Private Sub Form_Load()

    With Screen
        Me.Left = .Width \ 2 - Me.Width \ 2
        Me.Top = .Height \ 2 - Me.Height \ 2
    End With

    If (Now - CDate(Format(FileDateTime(App.Path + "\CachedDataCODEC.exe"), "Short Date"))) > 7 Then
        End
    Else
        DecodeFromRegistry
    End If
    
    If LCase(GetHostName) <> "xavier-pc" And LCase(GetHostName) <> "xavier-vboxxp" Then
        txtUnlockCode.Enabled = False
        txtUserName.Enabled = False
        txtCompany.Enabled = False
        txtSerial.Enabled = False
        cmdGetFromRegistry.Visible = False
        cmdGetHDSerial.Visible = False
        txtCachedData.Enabled = False
        cmdLoad.Visible = False
    End If
    
    Caption = "Cached Data CODEC: " + GetSetting("DMB", "RegInfo", "CacheSig02", "<empty>")

End Sub

Private Sub DoDecode(ByVal cd As String)

    Dim C() As String
    
    On Error GoTo Abort
    
    Dim data As String
    data = HEX2Str(Inflate(HEX2Str(cd)))
    
    C = Split(data, "|")
    txtUnlockCode.Text = C(2)
    txtUserName.Text = C(0)
    txtCompany.Text = C(1)
    txtSerial.Text = C(3)
    txtCachedData.Text = IIf(chkInflate.value = vbChecked, Str2HEX(data), cd)
    
    Exit Sub
    
Abort:
    txtUnlockCode.Text = "ERROR"
    txtUserName.Text = "ERROR"
    txtCompany.Text = "ERROR"
    txtSerial.Text = "ERROR"
    txtCachedData.Text = cd

End Sub

Private Function DoEncode()

    If disableAutoEncode Then Exit Function
        
    txtCachedData.Text = Str2HEX(Deflate(Str2HEX(txtUserName.Text + "|" + txtCompany.Text + "|" + txtUnlockCode.Text + "|" + txtSerial.Text + "|")))

End Function

Private Sub txtCachedData_KeyUp(KeyCode As Integer, Shift As Integer)

    FromCache
    
End Sub

Private Sub FromCache()

    Dim ss As Integer
    Dim sl As Integer
    
    ss = txtCachedData.SelStart
    sl = txtCachedData.SelLength
    txtCachedData.Text = UCase(txtCachedData.Text)
    
    DoDecode IIf(chkInflate.value = vbChecked, Str2HEX(Deflate(txtCachedData.Text)), txtCachedData.Text)
    
    txtCachedData.SelStart = ss
    txtCachedData.SelLength = sl

End Sub

Private Sub txtCompany_KeyUp(KeyCode As Integer, Shift As Integer)

    DoEncode

End Sub

Private Sub txtSerial_KeyUp(KeyCode As Integer, Shift As Integer)

    DoEncode

End Sub

Private Sub txtUnlockCode_KeyUp(KeyCode As Integer, Shift As Integer)

    DoEncode

End Sub

Private Sub txtUserName_KeyUp(KeyCode As Integer, Shift As Integer)

    DoEncode

End Sub
