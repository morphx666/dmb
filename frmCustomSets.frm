VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCustomSets 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manage Custom Sets"
   ClientHeight    =   3165
   ClientLeft      =   3750
   ClientTop       =   6795
   ClientWidth     =   6750
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
   ScaleHeight     =   3165
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add..."
      Height          =   360
      Left            =   5385
      TabIndex        =   5
      Top             =   375
      Width           =   1290
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "De&lete"
      Enabled         =   0   'False
      Height          =   360
      Left            =   5385
      TabIndex        =   4
      Top             =   2070
      Width           =   1290
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   5385
      TabIndex        =   3
      Top             =   2685
      Width           =   1290
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit..."
      Enabled         =   0   'False
      Height          =   360
      Left            =   5385
      TabIndex        =   2
      Top             =   810
      Width           =   1290
   End
   Begin MSComctlLib.ListView lvCustomSets 
      Height          =   2055
      Left            =   105
      TabIndex        =   1
      Top             =   375
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   3625
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Applies To"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Custom Sets"
      Height          =   195
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   900
   End
End
Attribute VB_Name = "frmCustomSets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()

    Dim i As Integer
    Dim idx As Integer

    ReDim Preserve CustomSets(UBound(CustomSets) + 1)
    With CustomSets(UBound(CustomSets))
        .Name = "NewCustomSet"
        idx = 1
ReStart:
        For i = 1 To UBound(CustomSets)
            If CustomSets(i).Name = .Name & Format(idx, "00") Then
                idx = idx + 1
                GoTo ReStart
            End If
        Next i
        .Name = .Name & Format(idx, "00")
        .AppliesTo = atcGroupsAndCommands
        
        ReDim .Alignment(0)
        ReDim .ColorBack(0)
        ReDim .ColorLine(0)
        ReDim .ColorNormalBack(0)
        ReDim .ColorNormalText(0)
        ReDim .ColorOverBack(0)
        ReDim .ColorOverText(0)
        ReDim .Cursor(0)
        ReDim .ImageBackNormal(0)
        ReDim .ImageBackOver(0)
        ReDim .ImageLeftNormal(0)
        ReDim .ImageLeftOver(0)
        ReDim .ImageRightNormal(0)
        ReDim .ImageRightOver(0)
        ReDim .nFontBold(0)
        ReDim .nFontItalic(0)
        ReDim .nFontName(0)
        ReDim .nFontSize(0)
        ReDim .oFontBold(0)
        ReDim .oFontItalic(0)
        ReDim .oFontName(0)
        ReDim .oFontSize(0)
        ReDim .SeparatorLength(0)
        ReDim .TargetFrame(0)
    End With
    SelCustomSet = UBound(CustomSets)

    frmCustomSetsDefine.Show vbModal
    
    InitDlg

End Sub

Private Sub cmdDelete_Click()

    Dim i As Integer

End Sub

Private Sub cmdEdit_Click()

    SelCustomSet = Val(lvCustomSets.SelectedItem.Tag)
    frmCustomSetsDefine.Show vbModal
    
    InitDlg

End Sub

Private Sub cmdOK_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    CenterForm Me
    SetupCharset Me
    
    InitDlg

End Sub

Private Sub InitDlg()

    Dim i As Integer
    
    lvCustomSets.ListItems.Clear
    For i = 1 To UBound(CustomSets)
        With lvCustomSets.ListItems.Add(, , CustomSets(i).Name)
            .Tag = i
            .SubItems(1) = GetAppliesToStr(CustomSets(i).AppliesTo)
        End With
    Next i
    
    If lvCustomSets.Visible Then lvCustomSets.SetFocus
    CoolListView lvCustomSets

End Sub

Private Function GetAppliesToStr(ByVal idx As Integer) As String

    Select Case idx
        Case atcCommands: GetAppliesToStr = "Commands"
        Case atcGroups: GetAppliesToStr = "Groups"
        Case atcGroupsAndCommands: GetAppliesToStr = "Groups and Commands"
        Case atcSeparators: GetAppliesToStr = "Separators"
    End Select

End Function

Private Sub lvCustomSets_ItemClick(ByVal Item As MSComctlLib.ListItem)

    cmdEdit.Enabled = True
    cmdDelete.Enabled = True

End Sub
