VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{9A2947B4-87EE-40EE-A3EF-32BDC32D5726}#1.0#0"; "xfxline3d.ocx"
Begin VB.Form frmCustomSetsDefine 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Define Custom Set"
   ClientHeight    =   6210
   ClientLeft      =   6090
   ClientTop       =   4875
   ClientWidth     =   6435
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
   ScaleHeight     =   6210
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   Begin xfxLine3D.ucLine3D uc3DLine1 
      Height          =   30
      Left            =   60
      TabIndex        =   14
      Top             =   5640
      Width           =   6345
      _ExtentX        =   11192
      _ExtentY        =   53
   End
   Begin VB.CommandButton cmdSelOp 
      Caption         =   "Inverse"
      Height          =   390
      Index           =   2
      Left            =   5235
      TabIndex        =   13
      Top             =   2640
      Width           =   1125
   End
   Begin VB.CommandButton cmdSelOp 
      Caption         =   "Select None"
      Height          =   390
      Index           =   1
      Left            =   5235
      TabIndex        =   12
      Top             =   2205
      Width           =   1125
   End
   Begin VB.CommandButton cmdSelOp 
      Caption         =   "Select All"
      Height          =   390
      Index           =   0
      Left            =   5235
      TabIndex        =   11
      Top             =   1770
      Width           =   1125
   End
   Begin VB.PictureBox picItemIcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   5235
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   690
      Visible         =   0   'False
      Width           =   240
   End
   Begin MSComctlLib.ImageList ilIcons 
      Left            =   4320
      Top             =   495
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lvValues 
      Height          =   3720
      Left            =   2730
      TabIndex        =   8
      Top             =   1770
      Width           =   2430
      _ExtentX        =   4286
      _ExtentY        =   6562
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ilIcons"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Value"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   345
      Left            =   3975
      TabIndex        =   7
      Top             =   5790
      Width           =   1125
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   5235
      TabIndex        =   6
      Top             =   5790
      Width           =   1125
   End
   Begin VB.ComboBox cmbAppliesTo 
      Height          =   315
      Left            =   75
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   975
      Width           =   2280
   End
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   75
      TabIndex        =   1
      Text            =   "NewCustomSet01"
      Top             =   315
      Width           =   3315
   End
   Begin MSComctlLib.TreeView tvProperties 
      Height          =   3720
      Left            =   75
      TabIndex        =   2
      Top             =   1770
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   6562
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
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
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Possible Values"
      Height          =   195
      Left            =   2730
      TabIndex        =   9
      Top             =   1530
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Applies To"
      Height          =   195
      Left            =   75
      TabIndex        =   4
      Top             =   735
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Properties"
      Height          =   195
      Left            =   75
      TabIndex        =   3
      Top             =   1530
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Set Title"
      Height          =   195
      Left            =   75
      TabIndex        =   0
      Top             =   90
      Width           =   585
   End
End
Attribute VB_Name = "frmCustomSetsDefine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum ResourceTypeConstants
    rtcIcon = &HF000
    rtcNone = 0
    rtcImage = 1
    rtcColor = 2
End Enum

Private CustomSetBack As CustomSetsDef
Private IsUpdating As Boolean

Private Sub cmbAppliesTo_Click()

    Dim nNode As Node

    CreateNodes
    
    IsUpdating = True
    
    lvValues.Visible = False
    tvProperties.Visible = False
    For Each nNode In tvProperties.Nodes
        tvProperties_NodeClick nNode
    Next nNode
    For Each nNode In tvProperties.Nodes
        nNode.Expanded = False
    Next nNode
    tvProperties_NodeClick tvProperties.Nodes(1)
    tvProperties.Visible = True
    lvValues.Visible = True
    
    IsUpdating = False

End Sub

Private Sub CreateNodes()

    Dim nNode As Node
    
    lvValues.ListItems.Clear
    
    With tvProperties.Nodes
        .Clear
        Set nNode = .Add(, tvwChild, "Color", GetLocalizedStr(212), IconIndex("Color")):
            Set nNode = .Add("Color", tvwChild, "ColorNormal", GetLocalizedStr(179), IconIndex("Normal")):
                Set nNode = .Add("ColorNormal", tvwChild, "ColorNormalText", GetLocalizedStr(196), IconIndex("Text Color")):
                Set nNode = .Add("ColorNormal", tvwChild, "ColorNormalBack", GetLocalizedStr(508), IconIndex("Back Color")):
            Set nNode = .Add("Color", tvwChild, "ColorOver", GetLocalizedStr(105), IconIndex("Over")):
                Set nNode = .Add("ColorOver", tvwChild, "ColorOverText", GetLocalizedStr(196), IconIndex("Text Color")):
                Set nNode = .Add("ColorOver", tvwChild, "ColorOverBack", GetLocalizedStr(508), IconIndex("Back Color")):
            Set nNode = .Add("Color", tvwChild, "ColorLine", GetLocalizedStr(509), IconIndex("Line")):
            Set nNode = .Add("Color", tvwChild, "ColorBack", GetLocalizedStr(508), IconIndex("Back Color")):
        Set nNode = .Add(, tvwChild, "SeparatorLength", GetLocalizedStr(791), IconIndex("SeparatorLength")):
        Set nNode = .Add(, tvwChild, "Font", GetLocalizedStr(213), IconIndex("Font")):
            Set nNode = .Add("Font", tvwChild, "FontNormal", GetLocalizedStr(179), IconIndex("Normal")):
                Set nNode = .Add("FontNormal", tvwChild, "nFontName", GetLocalizedStr(409), IconIndex("Font Name")):
                Set nNode = .Add("FontNormal", tvwChild, "nFontBold", GetLocalizedStr(510), IconIndex("Font Bold")):
                Set nNode = .Add("FontNormal", tvwChild, "nFontItalic", GetLocalizedStr(511), IconIndex("Font Italic")):
                Set nNode = .Add("FontNormal", tvwChild, "nFontUnderline", GetLocalizedStr(512), IconIndex("Font Underline")):
                Set nNode = .Add("FontNormal", tvwChild, "nFontSize", GetLocalizedStr(203), IconIndex("Size")):
            Set nNode = .Add("Font", tvwChild, "FontOver", GetLocalizedStr(105), IconIndex("Over")):
                Set nNode = .Add("FontOver", tvwChild, "oFontName", GetLocalizedStr(409), IconIndex("Font Name")):
                Set nNode = .Add("FontOver", tvwChild, "oFontBold", GetLocalizedStr(510), IconIndex("Font Bold")):
                Set nNode = .Add("FontOver", tvwChild, "oFontItalic", GetLocalizedStr(511), IconIndex("Font Italic")):
                Set nNode = .Add("FontOver", tvwChild, "oFontUnderline", GetLocalizedStr(512), IconIndex("Font Underline")):
                Set nNode = .Add("FontOver", tvwChild, "oFontSize", GetLocalizedStr(203), IconIndex("Size")):
            Set nNode = .Add("Font", tvwChild, "Alignment", GetLocalizedStr(115), IconIndex("Caption Alignment")):
        Set nNode = .Add(, tvwChild, "Cursor", GetLocalizedStr(215), IconIndex("Cursor")):
        Set nNode = .Add(, tvwChild, "Image", GetLocalizedStr(214), IconIndex("Image")):
            Set nNode = .Add("Image", tvwChild, "ImageLeft", GetLocalizedStr(190), IconIndex("Left")):
                Set nNode = .Add("ImageLeft", tvwChild, "ImageLeftNormal", GetLocalizedStr(179), IconIndex("Normal")):
                Set nNode = .Add("ImageLeft", tvwChild, "ImageLeftOver", GetLocalizedStr(105), IconIndex("Over")):
            Set nNode = .Add("Image", tvwChild, "ImageRight", GetLocalizedStr(191), IconIndex("Right")):
                Set nNode = .Add("ImageRight", tvwChild, "ImageRightNormal", GetLocalizedStr(179), IconIndex("Normal")):
                Set nNode = .Add("ImageRight", tvwChild, "ImageRightOver", GetLocalizedStr(105), IconIndex("Over")):
            Set nNode = .Add("Image", tvwChild, "ImageBack", GetLocalizedStr(513), IconIndex("Back Color")):
                Set nNode = .Add("ImageBack", tvwChild, "ImageBackNormal", GetLocalizedStr(179), IconIndex("Normal")):
                Set nNode = .Add("ImageBack", tvwChild, "ImageBackOver", GetLocalizedStr(105), IconIndex("Over")):
        
        Set nNode = .Add(, tvwChild, "Events", GetLocalizedStr(514), IconIndex("Events")):
            Set nNode = .Add("Events", tvwChild, "ActionType", GetLocalizedStr(108), IconIndex("Action Type")):
            Set nNode = .Add("Events", tvwChild, "TargetFrame", GetLocalizedStr(235), IconIndex("Target Frame")):
            
'        Set nNode = .Add(, tvwChild, "SFX", GetLocalizedStr(231), IconIndex("Special Effects")):
'            Set nNode = .Add("SFX", tvwChild, "SFXCommandHE", GetLocalizedStr(217), IconIndex("Highlight Effects")):
'                Set nNode = .Add("SFXCommandHE", tvwChild, "SFXCommandHEUseCBNormal", GetLocalizedStr(179), IconIndex("Normal")):
'                Set nNode = .Add("SFXCommandHE", tvwChild, "SFXCommandHEUseCBOver", GetLocalizedStr(105), IconIndex("Over")):
'                Set nNode = .Add("SFXCommandHE", tvwChild, "SFXCommandHEBorderSize", GetLocalizedStr(206), IconIndex("Border Size")):
'                Set nNode = .Add("SFXCommandHE", tvwChild, "SFXCommandHEMarginX", GetLocalizedStr(219), IconIndex("Command Horizontal Margin")):
'                Set nNode = .Add("SFXCommandHE", tvwChild, "SFXCommandHEMarginY", GetLocalizedStr(220), IconIndex("Command Vertical Margin")):
'            Set nNode = .Add("SFX", tvwChild, "SFXGroupSize", GetLocalizedStr(227), IconIndex("Size")):
'                Set nNode = .Add("SFXGroupSize", tvwChild, "SFXGroupWidth", GetLocalizedStr(516), IconIndex("Group Width")):
'                Set nNode = .Add("SFXGroupSize", tvwChild, "SFXGroupHeight", GetLocalizedStr(517), IconIndex("Group Height")):
'                Set nNode = .Add("SFXGroupSize", tvwChild, "SFXGroupScrolling", GetLocalizedStr(830), IconIndex("Scrolling")):
'            Set nNode = .Add("SFX", tvwChild, "SFXEx", GetLocalizedStr(218), IconIndex("Group Effects")):
'                Set nNode = .Add("SFXEx", tvwChild, "SFXDropShadow", GetLocalizedStr(831), IconIndex("Shadow")):
'                    Set nNode = .Add("SFXDropShadow", tvwChild, "SFXDropShadowSize", GetLocalizedStr(212), IconIndex("Size")):
'                    Set nNode = .Add("SFXDropShadow", tvwChild, "SFXDropShadowColor", GetLocalizedStr(203), IconIndex("Color")):                                                nNode.Tag = "Drop Shadow Color of the %IN% group"
'                Set nNode = .Add("SFXEx", tvwChild, "SFXTransparency", GetLocalizedStr(222), IconIndex("Transparency")):                                        nNode.Tag = "Transparency setting of the %IN% group"
        Set nNode = .Add(, tvwChild, "CommandsLayout", GetLocalizedStr(668), IconIndex("Commands Layout")):                                                nNode.Tag = "Commands Layout for the %IN% group"
        
        Select Case cmbAppliesTo.ListIndex
            Case 0
                .Remove "ColorLine"
                .Remove "SeparatorLength"
            Case 1
                .Remove "ColorLine"
                .Remove "SeparatorLength"
                .Remove "ColorBack"
                '.Remove "SFX"
                .Remove "CommandsLayout"
            Case 2
                .Remove "ColorLine"
                .Remove "SeparatorLength"
                .Remove "ColorBack"
                '.Remove "SFX"
                .Remove "CommandsLayout"
            Case 3
                .Remove "ColorNormal"
                .Remove "ColorOver"
                .Remove "Font"
                .Remove "Cursor"
                .Remove "Image"
                .Remove "Events"
                '.Remove "SFX"
                .Remove "CommandsLayout"
        End Select
    End With
    
End Sub

Private Sub cmdCancel_Click()

    CustomSets(SelCustomSet) = CustomSetBack
    
    Unload Me

End Sub

Private Sub cmdOK_Click()

    With CustomSets(SelCustomSet)
        .Name = txtTitle.Text
        .AppliesTo = cmbAppliesTo.ListIndex
    End With
    
    Unload Me

End Sub

Private Sub cmdSelOp_Click(Index As Integer)

    Dim nItem As ListItem
    
    For Each nItem In lvValues.ListItems
        Select Case Index
            Case 0: nItem.Checked = True
            Case 1: nItem.Checked = False
            Case 2: nItem.Checked = Not nItem.Checked
        End Select
        lvValues_ItemCheck nItem
    Next nItem

End Sub

Private Sub Form_Load()

    CustomSetBack = CustomSets(SelCustomSet)

    tvProperties.ImageList = frmMain.ilIcons

    CenterForm Me
    SetupCharset Me

    InitDialog

End Sub

Private Sub InitDialog()
    
    cmbAppliesTo.AddItem "Groups"
    cmbAppliesTo.AddItem "Commands"
    cmbAppliesTo.AddItem "Groups and Commands"
    cmbAppliesTo.AddItem "Separators"
    
    With CustomSets(SelCustomSet)
        txtTitle.Text = .Name
        cmbAppliesTo.ListIndex = .AppliesTo
    End With

End Sub

Private Sub lvValues_ItemCheck(ByVal Item As MSComctlLib.ListItem)

    Dim nItem As ListItem
    Dim tItem As Node
    Dim cItem As Node
    Dim SetBold As Boolean
    Dim SelValues As String
    Dim v() As String
    
    For Each nItem In lvValues.ListItems
        If nItem.Checked Then
            SetBold = True
            SelValues = SelValues + nItem.Text + "|"
        End If
    Next nItem
    
    Set tItem = tvProperties.SelectedItem
    Do
        If Not SetBold Then
            Set cItem = tItem.Child
            Do Until cItem Is Nothing
                If cItem.Bold Then GoTo SkipIt
                Set cItem = cItem.Next
            Loop
        End If
        tItem.Bold = SetBold
        Set tItem = tItem.Parent
    Loop Until tItem Is Nothing
    
SkipIt:
    
    If IsUpdating Then Exit Sub
    
    If SetBold Then
        If Right(SelValues, 1) = "|" Then SelValues = Left(SelValues, Len(SelValues) - 1)
        v = Split(SelValues, "|")
    Else
        ReDim v(0)
    End If
    
    Select Case tvProperties.SelectedItem.key
        Case "ColorNormalText"
            CustomSets(SelCustomSet).ColorNormalText = v
        Case "ColorNormalBack"
            CustomSets(SelCustomSet).ColorNormalBack = v
        Case "ColorOverText"
            CustomSets(SelCustomSet).ColorOverText = v
        Case "ColorOverBack"
            CustomSets(SelCustomSet).ColorOverBack = v
        Case "nFontName"
            CustomSets(SelCustomSet).nFontName = v
        Case "oFontName"
            CustomSets(SelCustomSet).oFontName = v
        Case "nFontBold"
            CustomSets(SelCustomSet).nFontBold = v
        Case "oFontBold"
            CustomSets(SelCustomSet).oFontBold = v
        Case "nFontItalic"
            CustomSets(SelCustomSet).nFontItalic = v
        Case "oFontItalic"
            CustomSets(SelCustomSet).oFontItalic = v
        Case "nFontUnderline"
            CustomSets(SelCustomSet).nFontUnderline = v
        Case "oFontUnderline"
            CustomSets(SelCustomSet).oFontUnderline = v
        Case "nFontSize"
            CustomSets(SelCustomSet).nFontSize = v
        Case "oFontSize"
            CustomSets(SelCustomSet).oFontSize = v
        Case "Alignment"
            CustomSets(SelCustomSet).Alignment = v
        Case "Cursor"
            CustomSets(SelCustomSet).Cursor = v
        Case "ActionType"
            CustomSets(SelCustomSet).ActionType = v
        Case "TargetFrame"
            CustomSets(SelCustomSet).TargetFrame = v
        Case "ColorLine"
            CustomSets(SelCustomSet).ColorLine = v
        Case "ColorBack"
            CustomSets(SelCustomSet).ColorBack = v
        Case "SeparatorLength"
            CustomSets(SelCustomSet).SeparatorLength = v
        Case "ImageLeftNormal"
            CustomSets(SelCustomSet).ImageLeftNormal = v
        Case "ImageLeftOver"
            CustomSets(SelCustomSet).ImageLeftOver = v
        Case "ImageRightNormal"
            CustomSets(SelCustomSet).ImageRightNormal = v
        Case "ImageRightOver"
            CustomSets(SelCustomSet).ImageRightOver = v
        Case "ImageBackNormal"
            CustomSets(SelCustomSet).ImageBackNormal = v
        Case "ImageBackOver"
            CustomSets(SelCustomSet).ImageBackOver = v
    End Select

End Sub

Private Sub tvProperties_NodeClick(ByVal Node As MSComctlLib.Node)

    Dim i As Integer
    Dim IsC As Boolean
    Dim IsG As Boolean
    Dim IsS As Boolean
    
    IsC = (cmbAppliesTo.ListIndex = 1) Or (cmbAppliesTo.ListIndex = 2)
    IsG = (cmbAppliesTo.ListIndex = 0) Or (cmbAppliesTo.ListIndex = 2)
    IsS = (cmbAppliesTo.ListIndex = 3)
    
    lvValues.ListItems.Clear
    Set lvValues.SmallIcons = Nothing
    ilIcons.ListImages.Clear
    ilIcons.ListImages.Add , , frmMain.Icon
    Set lvValues.SmallIcons = ilIcons
    
    Node.EnsureVisible
    Node.Selected = True

    Select Case Node.key
        Case "ColorNormalText"
            If IsC Then
                For i = 1 To UBound(MenuCmds)
                    AddValue GetRGB(MenuCmds(i).nTextColor, True), rtcColor, CustomSets(SelCustomSet).ColorNormalText
                Next i
            End If
            If IsG Then
                For i = 1 To UBound(MenuGrps)
                    AddValue GetRGB(MenuGrps(i).nTextColor, True), rtcColor, CustomSets(SelCustomSet).ColorNormalText
                Next i
            End If
        Case "ColorNormalBack"
            If IsC Then
                For i = 1 To UBound(MenuCmds)
                    AddValue GetRGB(MenuCmds(i).nBackColor, True), rtcColor, CustomSets(SelCustomSet).ColorNormalBack
                Next i
            End If
            If IsG Then
                For i = 1 To UBound(MenuGrps)
                    AddValue GetRGB(MenuGrps(i).nBackColor, True), rtcColor, CustomSets(SelCustomSet).ColorNormalBack
                Next i
            End If
        Case "ColorOverText"
            If IsC Then
                For i = 1 To UBound(MenuCmds)
                    AddValue GetRGB(MenuCmds(i).hTextColor, True), rtcColor, CustomSets(SelCustomSet).ColorOverText
                Next i
            End If
            If IsG Then
                For i = 1 To UBound(MenuGrps)
                    AddValue GetRGB(MenuGrps(i).hTextColor, True), rtcColor, CustomSets(SelCustomSet).ColorOverText
                Next i
            End If
        Case "ColorOverBack"
            If IsC Then
                For i = 1 To UBound(MenuCmds)
                    AddValue GetRGB(MenuCmds(i).hBackColor, True), rtcColor, CustomSets(SelCustomSet).ColorOverBack
                Next i
            End If
            If IsG Then
                For i = 1 To UBound(MenuGrps)
                    AddValue GetRGB(MenuGrps(i).hBackColor, True), rtcColor, CustomSets(SelCustomSet).ColorOverBack
                Next i
            End If
        Case "nFontName"
            If IsC Then
                For i = 1 To UBound(MenuCmds)
                    AddValue MenuCmds(i).NormalFont.FontName, rtcIcon + IconIndex("Font Name"), CustomSets(SelCustomSet).nFontName
                Next i
            End If
            If IsG Then
                For i = 1 To UBound(MenuGrps)
                    AddValue MenuGrps(i).DefNormalFont.FontName, rtcIcon + IconIndex("Font Name"), CustomSets(SelCustomSet).nFontName
                Next i
            End If
        Case "oFontName"
            If IsC Then
                For i = 1 To UBound(MenuCmds)
                    AddValue MenuCmds(i).NormalFont.FontName, rtcIcon + IconIndex("Font Name"), CustomSets(SelCustomSet).oFontName
                Next i
            End If
            If IsG Then
                For i = 1 To UBound(MenuGrps)
                    AddValue MenuGrps(i).DefNormalFont.FontName, rtcIcon + IconIndex("Font Name"), CustomSets(SelCustomSet).oFontName
                Next i
            End If
        Case "nFontBold"
            AddValue "Yes", rtcIcon + IconIndex("Font Bold"), CustomSets(SelCustomSet).nFontBold
        Case "oFontBold"
            AddValue "Yes", rtcIcon + IconIndex("Font Bold"), CustomSets(SelCustomSet).oFontBold
        Case "nFontItalic"
            AddValue "Yes", rtcIcon + IconIndex("Font Italic"), CustomSets(SelCustomSet).nFontItalic
        Case "oFontItalic"
            AddValue "Yes", rtcIcon + IconIndex("Font Italic"), CustomSets(SelCustomSet).oFontItalic
        Case "nFontUnderline"
            AddValue "Yes", rtcIcon + IconIndex("Font Underline"), CustomSets(SelCustomSet).nFontUnderline
        Case "oFontUnderline"
            AddValue "Yes", rtcIcon + IconIndex("Font Underline"), CustomSets(SelCustomSet).oFontUnderline
        Case "nFontSize"
            If IsC Then
                For i = 1 To UBound(MenuCmds)
                    AddValue MenuCmds(i).NormalFont.FontSize, rtcIcon + IconIndex("Size"), CustomSets(SelCustomSet).nFontSize
                Next i
            End If
            If IsG Then
                For i = 1 To UBound(MenuGrps)
                    AddValue MenuGrps(i).DefNormalFont.FontSize, rtcIcon + IconIndex("Size"), CustomSets(SelCustomSet).nFontSize
                Next i
            End If
        Case "oFontSize"
            If IsC Then
                For i = 1 To UBound(MenuCmds)
                    AddValue MenuCmds(i).HoverFont.FontSize, rtcIcon + IconIndex("Size"), CustomSets(SelCustomSet).oFontSize
                Next i
            End If
            If IsG Then
                For i = 1 To UBound(MenuGrps)
                    AddValue MenuGrps(i).DefHoverFont.FontSize, rtcIcon + IconIndex("Size"), CustomSets(SelCustomSet).oFontSize
                Next i
            End If
        Case "Alignment"
            AddValue "Left", rtcIcon + IconIndex("Left Align"), CustomSets(SelCustomSet).Alignment
            AddValue "Right", rtcIcon + IconIndex("Right Align"), CustomSets(SelCustomSet).Alignment
            AddValue "Center", rtcIcon + IconIndex("Center Align"), CustomSets(SelCustomSet).Alignment
        Case "Cursor"
            If IsC Then
                For i = 1 To UBound(MenuCmds)
                    AddValue GetCursorName(MenuCmds(i).iCursor), rtcIcon + IconIndex("Cursor"), CustomSets(SelCustomSet).Cursor
                Next i
            End If
            If IsG Then
                For i = 1 To UBound(MenuGrps)
                    AddValue GetCursorName(MenuGrps(i).iCursor), rtcIcon + IconIndex("Cursor"), CustomSets(SelCustomSet).Cursor
                Next i
            End If
        Case "ActionType"
            AddValue GetLocalizedStr(110), rtcIcon + IconIndex("NoEvents"), CustomSets(SelCustomSet).ActionType
            AddValue GetLocalizedStr(111), rtcIcon + IconIndex("URL"), CustomSets(SelCustomSet).ActionType
            AddValue GetLocalizedStr(112), rtcIcon + IconIndex("Group Alignment"), CustomSets(SelCustomSet).ActionType
            AddValue GetLocalizedStr(113), rtcIcon + IconIndex("New Window"), CustomSets(SelCustomSet).ActionType
        Case "TargetFrame"
            If IsC Then
                For i = 1 To UBound(MenuCmds)
                    With MenuCmds(i).Actions
                        AddValue .onclick.TargetFrame, rtcIcon + IconIndex("Target Frame"), CustomSets(SelCustomSet).TargetFrame, "(none defined)"
                        AddValue .OnDoubleClick.TargetFrame, rtcIcon + IconIndex("Target Frame"), CustomSets(SelCustomSet).TargetFrame, "(none defined)"
                        AddValue .onmouseover.TargetFrame, rtcIcon + IconIndex("Target Frame"), CustomSets(SelCustomSet).TargetFrame, "(none defined)"
                    End With
                Next i
            End If
            If IsG Then
                For i = 1 To UBound(MenuGrps)
                    With MenuGrps(i).Actions
                        AddValue .onclick.TargetFrame, rtcIcon + IconIndex("Target Frame"), CustomSets(SelCustomSet).TargetFrame, "(none defined)"
                        AddValue .OnDoubleClick.TargetFrame, rtcIcon + IconIndex("Target Frame"), CustomSets(SelCustomSet).TargetFrame, "(none defined)"
                        AddValue .onmouseover.TargetFrame, rtcIcon + IconIndex("Target Frame"), CustomSets(SelCustomSet).TargetFrame, "(none defined)"
                    End With
                Next i
            End If
        Case "ColorLine"
            For i = 1 To UBound(MenuCmds)
                If MenuCmds(i).Name = "[SEP]" Then
                    AddValue GetRGB(MenuCmds(i).nTextColor, True), rtcColor, CustomSets(SelCustomSet).ColorLine
                End If
            Next i
        Case "ColorBack"
            If IsG Then
                For i = 1 To UBound(MenuGrps)
                    AddValue GetRGB(MenuGrps(i).bColor, True), rtcColor, CustomSets(SelCustomSet).ColorBack
                Next i
            End If
            If IsS Then
                For i = 1 To UBound(MenuCmds)
                    If MenuCmds(i).Name = "[SEP]" Then
                        AddValue GetRGB(MenuCmds(i).nBackColor, True), rtcColor, CustomSets(SelCustomSet).ColorBack
                    End If
                Next i
            End If
        Case "SeparatorLength"
            For i = 1 To UBound(MenuCmds)
                If MenuCmds(i).Name = "[SEP]" Then
                    AddValue MenuCmds(i).SeparatorPercent & "%", rtcNone, CustomSets(SelCustomSet).SeparatorLength
                End If
            Next i
        Case "ImageLeftNormal"
            If IsC Then
                For i = 1 To UBound(MenuCmds)
                    AddValue MenuCmds(i).LeftImage.NormalImage, rtcImage, CustomSets(SelCustomSet).ImageLeftNormal, "(no image)"
                Next i
            End If
            If IsG Then
                For i = 1 To UBound(MenuGrps)
                    AddValue MenuGrps(i).tbiLeftImage.NormalImage, rtcImage, CustomSets(SelCustomSet).ImageLeftNormal, "(no image)"
                Next i
            End If
        Case "ImageLeftOver"
            If IsC Then
                For i = 1 To UBound(MenuCmds)
                    AddValue MenuCmds(i).LeftImage.HoverImage, rtcImage, CustomSets(SelCustomSet).ImageLeftOver, "(no image)"
                Next i
            End If
            If IsG Then
                For i = 1 To UBound(MenuGrps)
                    AddValue MenuGrps(i).tbiLeftImage.HoverImage, rtcImage, CustomSets(SelCustomSet).ImageLeftOver, "(no image)"
                Next i
            End If
        Case "ImageRightNormal"
            If IsC Then
                For i = 1 To UBound(MenuCmds)
                    AddValue MenuCmds(i).RightImage.NormalImage, rtcImage, CustomSets(SelCustomSet).ImageRightNormal, "(no image)"
                Next i
            End If
            If IsG Then
                For i = 1 To UBound(MenuGrps)
                    AddValue MenuGrps(i).tbiRightImage.NormalImage, rtcImage, CustomSets(SelCustomSet).ImageRightNormal, "(no image)"
                Next i
            End If
        Case "ImageRightOver"
            If IsC Then
                For i = 1 To UBound(MenuCmds)
                    AddValue MenuCmds(i).RightImage.HoverImage, rtcImage, CustomSets(SelCustomSet).ImageRightOver, "(no image)"
                Next i
            End If
            If IsG Then
                For i = 1 To UBound(MenuGrps)
                    AddValue MenuGrps(i).tbiRightImage.HoverImage, rtcImage, CustomSets(SelCustomSet).ImageRightOver, "(no image)"
                Next i
            End If
        Case "ImageBackNormal"
            If IsC Then
                For i = 1 To UBound(MenuCmds)
                    AddValue MenuCmds(i).BackImage.NormalImage, rtcImage, CustomSets(SelCustomSet).ImageBackNormal, "(no image)"
                Next i
            End If
            If IsG Then
                For i = 1 To UBound(MenuGrps)
                    AddValue MenuGrps(i).tbiBackImage.NormalImage, rtcImage, CustomSets(SelCustomSet).ImageBackNormal, "(no image)"
                Next i
            End If
        Case "ImageBackOver"
            If IsC Then
                For i = 1 To UBound(MenuCmds)
                    AddValue MenuCmds(i).BackImage.HoverImage, rtcImage, CustomSets(SelCustomSet).ImageLeftOver, "(no image)"
                Next i
            End If
            If IsG Then
                For i = 1 To UBound(MenuGrps)
                    AddValue MenuGrps(i).tbiBackImage.HoverImage, rtcImage, CustomSets(SelCustomSet).ImageBackOver, "(no image)"
                Next i
            End If
        Case Else
            'Debug.Print Node.key
    End Select

End Sub

Private Sub AddValue(ByVal v As String, ByVal ResType As ResourceTypeConstants, a() As String, Optional ByVal EmptySubs As String = "")
    
    Dim mv As String
    Dim nItem As ListItem
    
    mv = v
    If ResType = rtcImage Then v = GetFileName(v)
    
    If EmptySubs <> "" And v = "" Then v = EmptySubs
    Set nItem = lvValues.FindItem(v, lvwText, , lvwWhole)
    If nItem Is Nothing Then
        Set nItem = lvValues.ListItems.Add(, , v)
        If Not IsUpdating Then
            With nItem
                If (ResType And rtcIcon) = rtcIcon Then
                    ilIcons.ListImages.Add , , frmMain.ilIcons.ListImages(ResType Xor rtcIcon).Picture
                Else
                    picItemIcon.Picture = LoadPicture()
                    picItemIcon.Cls
                    Select Case ResType
                        Case rtcNone
                            ilIcons.ListImages.Add , , picItemIcon.Image
                        Case rtcColor
                            If v = "Transparent" Then
                                picItemIcon.Picture = frmMain.ilIcons.ListImages("Transparent").Picture
                            Else
                                picItemIcon.Line (0, 1)-(15, 15), Val(Replace(v, "#", "&h")), BF
                            End If
                            picItemIcon.Line (0, 1)-(15, 15), vbBlack, B
                            Set picItemIcon.Picture = picItemIcon.Image
                            ilIcons.ListImages.Add , , picItemIcon.Picture
                        Case rtcImage
                            Set picItemIcon.Picture = LoadPictureRes(mv)
                            Set picItemIcon.Picture = picItemIcon.Image
                            ilIcons.ListImages.Add , , picItemIcon.Picture
                    End Select
                End If
                .SmallIcon = ilIcons.ListImages.count
            End With
            CoolListView lvValues
        End If
        
        If (InStr("|" + Join(a, "|") + "|", "|" + v + "|") > 0) Then
            nItem.Checked = True
            If Not tvProperties.SelectedItem.Bold Then lvValues_ItemCheck nItem
        End If
    End If
    
End Sub

