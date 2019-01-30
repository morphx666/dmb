VERSION 5.00
Object = "{E1C6DB9D-BD4A-4A61-A759-0CED75D034BF}#43.0#0"; "SmartButton.ocx"
Begin VB.Form frmAUE 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Advanced Unfolding Effects"
   ClientHeight    =   4725
   ClientLeft      =   6330
   ClientTop       =   5265
   ClientWidth     =   5595
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAUE.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrFixDblClickBug 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   2775
      Top             =   4275
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview..."
      Height          =   375
      Left            =   30
      TabIndex        =   10
      Top             =   4305
      Width           =   1080
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3645
      TabIndex        =   11
      Top             =   4305
      Width           =   900
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4650
      TabIndex        =   12
      Top             =   4305
      Width           =   900
   End
   Begin VB.Frame Frame2 
      Caption         =   "DX Effects"
      Height          =   3135
      Left            =   30
      TabIndex        =   3
      Top             =   1095
      Width           =   2220
      Begin SmartButtonProject.SmartButton cmdAUE 
         Height          =   360
         Left            =   120
         TabIndex        =   5
         Top             =   2640
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   635
         Caption         =   "        Microsoft Help..."
         Picture         =   "frmAUE.frx":038A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionLayout   =   3
         PictureLayout   =   3
         OffsetLeft      =   4
      End
      Begin VB.ListBox lstEffects 
         Height          =   2310
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   4
         Top             =   255
         Width           =   1965
      End
   End
   Begin VB.Frame frameParams 
      Caption         =   "Effect Parameters"
      Height          =   3135
      Left            =   2340
      TabIndex        =   6
      Top             =   1095
      Width           =   3210
      Begin VB.TextBox txtNum 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   0
         Left            =   1185
         TabIndex        =   9
         Top             =   1005
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox cmbCombo 
         Height          =   315
         Index           =   0
         Left            =   1905
         TabIndex        =   8
         Top             =   345
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   7
         Top             =   330
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   375
      End
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   525
      Picture         =   "frmAUE.frx":0724
      Top             =   600
      Width           =   240
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AUE Engine developed by xFX JumpStart"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   975
      TabIndex        =   1
      Top             =   480
      Width           =   2520
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DirectX is a technology of Microsoft Corp."
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   975
      TabIndex        =   2
      Top             =   630
      Width           =   2625
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   75
      Left            =   0
      Top             =   885
      Width           =   5610
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "AUE Engine 1.0 for Internet Explorer's DirectX Transform Filters "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   975
      TabIndex        =   0
      Top             =   30
      Width           =   4020
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   105
      Picture         =   "frmAUE.frx":0AAE
      Top             =   75
      Width           =   720
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   885
      Left            =   0
      Top             =   0
      Width           =   5610
   End
End
Attribute VB_Name = "frmAUE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#If LITE = 0 Then

Private Type TEffects
    Name As String
    params() As String
    Values() As String
    Enabled As Boolean
End Type
Private Effects() As TEffects

Private IsUpdatting As Boolean

Private OriginalFilter As String

Private Sub cmbCombo_Change(Index As Integer)

    If IsUpdatting Then Exit Sub
    Effects(lstEffects.ListIndex + 1).Values(cmbCombo(Index).Tag) = cmbCombo(Index).Text

End Sub

Private Sub cmbCombo_Click(Index As Integer)

    If IsUpdatting Then Exit Sub
    Effects(lstEffects.ListIndex + 1).Values(cmbCombo(Index).Tag) = cmbCombo(Index).Text

End Sub

Private Sub cmdAUE_Click()

    RunShellExecute "Open", "http://msdn.microsoft.com/library/default.asp?url=/workshop/author/filter/reference/filters/" + LCase(lstEffects.List(lstEffects.ListIndex)) + ".asp", 0, 0, 0

End Sub

Private Sub cmdCancel_Click()

    Project.DXFilter = OriginalFilter
    Unload Me

End Sub

Private Sub cmdOK_Click()

    GenerateFilterCode
    Unload Me

End Sub

Private Sub cmdPreview_Click()

    GenerateFilterCode
    frmMain.ShowPreview

End Sub

Private Sub GenerateFilterCode()

    Dim i As Integer
    Dim j As Integer
    Dim fCode As String
    
    For i = 1 To UBound(Effects)
        With Effects(i)
            If .Enabled Then
                fCode = fCode + "progid:DXImageTransform.Microsoft." + .Name + "("
                For j = 1 To UBound(.params)
                    fCode = fCode + Split(.params(j), "|")(0) + "=" + .Values(j) + ","
                Next j
                fCode = Left(fCode, Len(fCode) - 1) + ")"
            End If
        End With
    Next i
    
    Project.DXFilter = fCode

End Sub

Private Sub Form_Load()

    CenterForm Me
    SetupCharset Me
    
    OriginalFilter = Project.DXFilter
    
    LoadEffects

End Sub

Private Sub LoadEffects()
    
    Dim lines() As String
    Dim i As Integer
    Dim j As Integer
    Dim en As Integer
    Dim pn As Integer
    
    ReDim Effects(0)
    
    lstEffects.Clear
    
    lines = Split(LoadFile(AppPath + "rsc\aue.dat"), vbCrLf)
    
    IsUpdatting = True
    
    For i = 0 To UBound(lines)
        If Left(lines(i), 1) = "#" Then
            en = UBound(Effects) + 1
            ReDim Preserve Effects(en)
            With Effects(en)
                .Name = Mid(lines(i), 2)
                lstEffects.AddItem .Name
                lstEffects.Selected(lstEffects.NewIndex) = (InStr(Project.DXFilter, "." + .Name + "(") > 0)
                .Enabled = lstEffects.Selected(lstEffects.NewIndex)
                ReDim .params(0)
                ReDim .Values(0)
                For j = i + 1 To UBound(lines)
                    If Left(lines(j), 1) = "#" Then
                        i = j - 1
                        Exit For
                    End If
                    pn = UBound(.params) + 1
                    ReDim Preserve .params(pn)
                    ReDim Preserve .Values(pn)
                    .params(pn) = Mid(lines(j), 2)
                    If .Enabled Then
                        .Values(pn) = ExtractParamValue(.Name, CStr(Split(.params(pn), "|")(0)))
                    End If
                    If LenB(.Values(pn)) = 0 Then
                        .Values(pn) = CStr(Split(.params(pn), "|")(2))
                    End If
                Next j
            End With
        End If
    Next i
    
    IsUpdatting = False
    
    lstEffects.ListIndex = 0
    lstEffects_Click
    
End Sub

Private Function ExtractParamValue(fName As String, pName As String) As String
    
    Dim f As String
    Dim p1 As Integer
    Dim p2 As Integer
    
    On Error GoTo ExitFcn
    
    p1 = InStr(Project.DXFilter, fName)
    p2 = InStr(p1, Project.DXFilter, ")")
    f = Mid(Project.DXFilter, p1, p2 - p1)
    
    p1 = InStr(f, pName)
    p2 = InStr(p1, f, "=")
    f = Mid(f, p2 + 1)
    
    p1 = InStr(f, ",") - 1
    If p1 = -1 Then p1 = Len(f)
    ExtractParamValue = Left(f, p1)
    
ExitFcn:
    
End Function

Private Sub lstEffects_Click()

    If IsUpdatting Then Exit Sub
    LoadControls
    
End Sub

Private Sub LoadControls()

    Dim i As Integer
    Dim k As Integer
    Dim p() As String
    Dim v() As String
    Dim e As Boolean
    Dim j As Integer
    
    Dim y As Long
    
    While cmbCombo.Count > 1
        Unload cmbCombo(cmbCombo.Count - 1)
    Wend
    
    While lblLabel.Count > 1
        Unload lblLabel(lblLabel.Count - 1)
    Wend
    
    While txtNum.Count > 1
        Unload txtNum(txtNum.Count - 1)
    Wend
    
    IsUpdatting = True
    
    y = lblLabel(0).Top
    e = lstEffects.Selected(lstEffects.ListIndex)
    p = Effects(lstEffects.ListIndex + 1).params
    v = Effects(lstEffects.ListIndex + 1).Values
    
    Effects(lstEffects.ListIndex + 1).Enabled = e
    
    For i = 1 To UBound(p)
        k = lblLabel.Count
        Load lblLabel(k)
        With lblLabel(k)
            .Caption = StrConv(Split(p(i), "|")(0), vbProperCase)
            .Move lblLabel(0).Top, y
            .Enabled = e
            .Visible = True
        End With
        Select Case CStr(Split(p(i), "|")(1))
            Case "n"
                k = txtNum.Count
                Load txtNum(k)
                With txtNum(k)
                    .Text = v(i)
                    .Move frameParams.Width / 2, y - (.Height - lblLabel(0).Height) \ 2
                    .Enabled = e
                    .Tag = i
                    .Visible = True
                End With
            Case "c"
                k = cmbCombo.Count
                Load cmbCombo(k)
                With cmbCombo(k)
                    .Text = v(i)
                    For j = 2 To UBound(Split(p(i), "|"))
                        .AddItem Split(p(i), "|")(j)
                    Next j
                    .Move frameParams.Width / 2, y - (.Height - lblLabel(0).Height) \ 2
                    .Enabled = e
                    .Tag = i
                    .Visible = True
                End With
        End Select
        y = y + (lblLabel(0).Height + 135)
    Next i
    
    With frameParams
        .Caption = lstEffects.List(lstEffects.ListIndex) + " Parameters"
        .Enabled = e
    End With
    
    IsUpdatting = False

End Sub

Private Sub lstEffects_DblClick()

    tmrFixDblClickBug.Enabled = True

End Sub

Private Sub tmrFixDblClickBug_Timer()

    tmrFixDblClickBug.Enabled = False
    lstEffects_Click

End Sub

Private Sub txtNum_Change(Index As Integer)

    If IsUpdatting Then Exit Sub
    Effects(lstEffects.ListIndex + 1).Values(txtNum(Index).Tag) = txtNum(Index).Text

End Sub

Private Sub txtNum_GotFocus(Index As Integer)

    SelAll txtNum(Index)

End Sub

#End If
