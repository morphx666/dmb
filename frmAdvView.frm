VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAdvView 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Item Properties"
   ClientHeight    =   3000
   ClientLeft      =   7785
   ClientTop       =   5955
   ClientWidth     =   4620
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
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
   ScaleHeight     =   3000
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdBrowse 
      Height          =   315
      Left            =   2595
      Picture         =   "frmAdvView.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2415
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.CommandButton cmdWinParams 
      Height          =   315
      Left            =   3300
      Picture         =   "frmAdvView.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "New Window Parameters"
      Top             =   2340
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.CommandButton cmdFont 
      Height          =   315
      Left            =   3900
      Picture         =   "frmAdvView.frx":048C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2100
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.CommandButton cmdColor 
      BackColor       =   &H80000007&
      Height          =   315
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1530
      Visible         =   0   'False
      WhatsThisHelpID =   20120
      Width           =   315
   End
   Begin VB.TextBox txtText 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2670
      TabIndex        =   2
      Top             =   990
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.ComboBox cmbCombo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmAdvView.frx":05D6
      Left            =   2655
      List            =   "frmAdvView.frx":05D8
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   495
      Visible         =   0   'False
      Width           =   1395
   End
   Begin MSComctlLib.ImageList ilIcons 
      Left            =   195
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdvView.frx":05DA
            Key             =   "Color"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdvView.frx":0736
            Key             =   "Font"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdvView.frx":0892
            Key             =   "Cursor"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdvView.frx":09EE
            Key             =   "Image"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdvView.frx":0B4A
            Key             =   "Special Effects"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdvView.frx":0CA6
            Key             =   "General"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdvView.frx":0DBE
            Key             =   "Margins"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdvView.frx":1158
            Key             =   "Sound"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdvView.frx":12B2
            Key             =   "Action: Over"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdvView.frx":140C
            Key             =   "Action: Click"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdvView.frx":1566
            Key             =   "Action: Double Click"
         EndProperty
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid mfgData 
      Height          =   2010
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3180
      _ExtentX        =   5609
      _ExtentY        =   3545
      _Version        =   393216
      Rows            =   6
      Cols            =   6
      BackColor       =   16777215
      GridColorFixed  =   8421504
      HighLight       =   0
      GridLinesFixed  =   1
      ScrollBars      =   2
      Appearance      =   0
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
End
Attribute VB_Name = "frmAdvView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sCtrl As Control
Dim sCtrl2 As Control
Dim IsUpdating As Boolean

Private Const Section_DisColor = &H808080
Private Const Property_DisColor = &HE0E0E0

Private Enum CellTypesConstants
    ctcText = 0
    ctcImage = 1
    ctcColor = 2
    ctcYesNo = 3
    ctcNumber = 4
    ctcFile = 5
    ctcNewWindowParams = 6
    ctcCursorList = 7
    ctcFont = 8
    ctcActionTypeList = 9
    ctcTargetMenuList = 10
    ctcTargetFrameList = 11
    ctcAlignmentList = 12
    ctcGroupAlignmentList = 13
End Enum

Private Function GetGroupAlignment(idx As Integer) As String

    Select Case idx
        Case 0
            GetGroupAlignment = "Bottom / Left"
        Case 1
            GetGroupAlignment = "Bottom / Right"
        Case 2
            GetGroupAlignment = "Top / Left"
        Case 3
            GetGroupAlignment = "Top / Right"
        Case 4
            GetGroupAlignment = "Left / Top"
        Case 5
            GetGroupAlignment = "Left / Bottom"
        Case 6
            GetGroupAlignment = "Right / Top"
        Case 7
            GetGroupAlignment = "Right / Bottom"
    End Select

End Function

Private Sub Form_Load()

    InitGrid
    Update

End Sub

Private Sub Form_Resize()

    On Error Resume Next

    mfgData.Move 0, 0, Width - 150, Height - 380
    
    With mfgData
        .ColWidth(0) = 1800
        .ColWidth(1) = (mfgData.Width - .ColWidth(0) - 350 - 15) / 2
        .ColWidth(2) = 350
        .ColWidth(3) = .ColWidth(1)
    End With
    
    mfgData_Scroll

End Sub

Friend Sub Update()

    Dim IsG As Boolean
    Dim IsC As Boolean
    Dim IsS As Boolean
    
    If frmMain.tvMenus.SelectedItem Is Nothing Then Exit Sub
    
    With frmMain.tvMenus.SelectedItem
        IsG = IsGroup(.key)
        IsC = IsCommand(.key)
        IsS = IsSeparator(.key)
    End With
    
    If IsC Then
        With MenuCmds(GetID)
            SetValue "General", "Name", .Name, , True
            SetValue "General", "Caption", .Caption, , True
            SetValue "General", "Enabled", IIf(Not .disabled, "Yes", "No"), , True
            SetValue "General", "Status Text", .WinStatus, , True
            
            SetValue "Color", "Text", GetRGB(.nTextColor), GetRGB(.hTextColor), True, True
            SetValue "Color", "Background", GetRGB(.nBackColor), GetRGB(.hBackColor), True, True
            
            SetValue "Cursor", "Type", , StrConv(GetCursorName(.iCursor), vbProperCase), False, True
            
            SetValue "Font", "Name", .NormalFont.FontName, .HoverFont.FontName, True, True
            SetValue "Font", "Size", CStr(.NormalFont.FontSize), CStr(.HoverFont.FontSize), True, True
            SetValue "Font", "Bold", IIf(.NormalFont.FontBold, "Yes", "No"), IIf(.HoverFont.FontBold, "Yes", "No"), True, True
            SetValue "Font", "Italic", IIf(.NormalFont.FontItalic, "Yes", "No"), IIf(.HoverFont.FontItalic, "Yes", "No"), True, True
            SetValue "Font", "Underline", IIf(.NormalFont.FontUnderline, "Yes", "No"), IIf(.HoverFont.FontUnderline, "Yes", "No"), True, True
            SetValue "Font", "Alignment", StrConv(GetAlignmentName(.Alignment), vbProperCase), , True, False
            
            SetValue "Image", "Left", .LeftImage.NormalImage, .LeftImage.HoverImage, True, True
            SetValue "Image", "Left Width", CStr(.LeftImage.w), , True
            SetValue "Image", "Left Height", CStr(.LeftImage.h), , True
            SetValue "Image", "Right", .RightImage.NormalImage, .RightImage.HoverImage, True, True
            SetValue "Image", "Right Width", CStr(.RightImage.w), , True
            SetValue "Image", "Right Height", CStr(.RightImage.h), , True
            SetValue "Image", "Background", .BackImage.NormalImage, .BackImage.HoverImage, True, True
            
            UpdateAction "Over", .Actions.onmouseover
            UpdateAction "Click", .Actions.onclick
            UpdateAction "Double Click", .Actions.OnDoubleClick
        End With
    End If
    
    If IsG Then
        With MenuGrps(GetID)
            SetValue "General", "Name", .Name, , True, False
            SetValue "General", "Caption", .Caption, , True, False
            SetValue "General", "Enabled", IIf(Not .disabled, "Yes", "No"), , True, False
            SetValue "General", "Status Text", .WinStatus, , True, False
            
            SetValue "Color", "Text", GetRGB(.nTextColor), GetRGB(.hTextColor), True, True
            SetValue "Color", "Background", GetRGB(.nBackColor), GetRGB(.hBackColor), True, True
            
            SetValue "Cursor", "Type", , StrConv(GetCursorName(.iCursor), vbProperCase), False, True
            
            SetValue "Font", "Name", .DefNormalFont.FontName, .DefHoverFont.FontName, True, True
            SetValue "Font", "Size", CStr(.DefNormalFont.FontSize), CStr(.DefHoverFont.FontSize), True, True
            SetValue "Font", "Bold", IIf(.DefNormalFont.FontBold, "Yes", "No"), IIf(.DefHoverFont.FontBold, "Yes", "No"), True, True
            SetValue "Font", "Italic", IIf(.DefNormalFont.FontItalic, "Yes", "No"), IIf(.DefHoverFont.FontItalic, "Yes", "No"), True, True
            SetValue "Font", "Underline", IIf(.DefNormalFont.FontUnderline, "Yes", "No"), IIf(.DefHoverFont.FontUnderline, "Yes", "No"), True, True
            SetValue "Font", "Alignment", StrConv(GetAlignmentName(.CaptionAlignment), vbProperCase), , True, False
            
            SetValue "Image", "Left", .LeftImage.NormalImage, .LeftImage.HoverImage, True, True
            SetValue "Image", "Left Width", CStr(.LeftImage.w), , True, False
            SetValue "Image", "Left Height", CStr(.LeftImage.h), , True, False
            SetValue "Image", "Right", .RightImage.NormalImage, .RightImage.HoverImage, True, True
            SetValue "Image", "Right Width", CStr(.RightImage.w), , True, False
            SetValue "Image", "Right Height", CStr(.RightImage.h), , True, False
            SetValue "Image", "Background", .BackImage.NormalImage, .BackImage.HoverImage, True, True
            
            UpdateAction "Over", .Actions.onmouseover
            UpdateAction "Click", .Actions.onclick
            UpdateAction "Double Click", .Actions.OnDoubleClick
        End With
    End If

End Sub

Private Sub UpdateAction(ActionName As String, sAction As Action)

    With sAction
        SetValue "Action: " + ActionName, "Type", GetActionName(.Type), , True, False
        SetValue "Action: " + ActionName, "URL / Script", .URL, , .Type = atcURL Or .Type = atcNewWindow, False
        SetValue "Action: " + ActionName, "Target Menu", MenuGrps(.TargetMenu).Name, , .Type = atcCascade, False
        SetValue "Action: " + ActionName, "Target Frame", .TargetFrame, , .Type = atcURL, False
        SetValue "Action: " + ActionName, "Submenu Alignment", GetGroupAlignment(CInt(.TargetMenuAlignment)), , .Type = atcCascade, False
        SetValue "Action: " + ActionName, "New Window Params", DecodeNWP(.URL, .WindowOpenParams), , .Type = atcNewWindow, False
    End With

End Sub

Private Function GetActionName(ActionIdx As ActionTypeConstants) As String

    Select Case ActionIdx
        Case atcCascade
            GetActionName = "Display SubMenu"
        Case atcNewWindow
            GetActionName = "Open URL in new Window"
        Case atcNone
            GetActionName = "(none)"
        Case atcURL
            GetActionName = "Open URL / Execute Script"
    End Select

End Function

Private Sub SetValue(Section As String, Property As String, Optional Normal As String, Optional Over As String, Optional nState As Boolean, Optional oState As Boolean)

    Dim r As Long
    Dim rp As Long
    
    IsUpdating = True

Retry:
    With mfgData
        For r = 0 To .rows - 1
            If .TextMatrix(r, 0) = Section Then
                For rp = r + 1 To .rows - 1
                    If Property = Trim(.TextMatrix(rp, 0)) Then
                        .TextMatrix(rp, 1) = IIf(nState, Normal, "")
                        .TextMatrix(rp, 3) = IIf(oState, Over, "")
                        .Col = 1: .Row = rp
                        .CellBackColor = IIf(nState, vbWhite, Property_DisColor)
                        .Col = 3: .Row = rp
                        .CellBackColor = IIf(oState, vbWhite, Property_DisColor)
                        Exit For
                    End If
                Next rp
                Exit For
            End If
        Next r
    End With
    
    IsUpdating = False

End Sub

Private Sub InitGrid()

    IsUpdating = True

    With mfgData
        .Clear
        .cols = 4
        .rows = 1
        
        .FillStyle = flexFillRepeat
        .ScrollBars = flexScrollBarVertical
        .HighLight = flexHighlightNever
        .FocusRect = flexFocusNone
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeNone
        .AllowBigSelection = False
        .SelectionMode = flexSelectionFree
        .GridLines = flexGridFlat
        .TextStyle = flexTextFlat
        .MergeCells = flexMergeNever
        .GridLinesFixed = flexGridFlat
        .TextStyleFixed = flexTextFlat
        .PictureType = flexPictureColor
        
        .Row = 0: .Col = 1
        .Text = "Normal"
        .CellAlignment = flexAlignCenterCenter
        .CellFontBold = True
        
        .Row = 0: .Col = 3
        .Text = "Over"
        .CellAlignment = flexAlignCenterCenter
        .CellFontBold = True
        
        .Row = 0: .Col = 1
        .RowSel = 0: .ColSel = 3
        .CellAlignment = flexAlignCenterCenter
        
        AddSection "General"
        AddProperty "General", "Name", ctcText, , True
        AddProperty "General", "Caption", ctcText, , True
        AddProperty "General", "Enabled", ctcYesNo, , True
        AddProperty "General", "Status Text", ctcText, , True
        
        AddSection "Color"
        AddProperty "Color", "Text", ctcColor
        AddProperty "Color", "Background", ctcColor
        
        AddSection "Cursor"
        AddProperty "Cursor", "Type", ctcCursorList, True
        
        AddSection "Font"
        AddProperty "Font", "Name", ctcFont
        AddProperty "Font", "Size", ctcNumber
        AddProperty "Font", "Bold", ctcYesNo
        AddProperty "Font", "Italic", ctcYesNo
        AddProperty "Font", "Underline", ctcYesNo
        AddProperty "Font", "Alignment", ctcAlignmentList, , True
        
        AddSection "Image"
        AddProperty "Image", "Left", ctcFile
        AddProperty "Image", "Left Width", ctcNumber, , True
        AddProperty "Image", "Left Height", ctcNumber, , True
        AddProperty "Image", "Right", ctcFile
        AddProperty "Image", "Right Width", ctcNumber, , True
        AddProperty "Image", "Right Height", ctcNumber, , True
        AddProperty "Image", "Background", ctcFile
        
        AddSection "Action: Over"
        AddProperty "Action: Over", "Type", ctcActionTypeList, , True
        AddProperty "Action: Over", "URL / Script", ctcFile, , True
        AddProperty "Action: Over", "Target Menu", ctcTargetMenuList, , True
        AddProperty "Action: Over", "Target Frame", ctcTargetFrameList, , True
        AddProperty "Action: Over", "Submenu Alignment", ctcGroupAlignmentList, , True
        AddProperty "Action: Over", "New Window Params", ctcNewWindowParams, , True
        
        AddSection "Action: Click"
        AddProperty "Action: Click", "Type", ctcActionTypeList, , True
        AddProperty "Action: Click", "URL / Script", ctcFile, , True
        AddProperty "Action: Click", "Target Menu", ctcTargetMenuList, , True
        AddProperty "Action: Click", "Target Frame", ctcTargetFrameList, , True
        AddProperty "Action: Click", "Submenu Alignment", ctcGroupAlignmentList, , True
        AddProperty "Action: Click", "New Window Params", ctcNewWindowParams, , True
        
        AddSection "Action: Double Click"
        AddProperty "Action: Double Click", "Type", ctcActionTypeList, , True
        AddProperty "Action: Double Click", "URL / Script", ctcFile, , True
        AddProperty "Action: Double Click", "Target Menu", ctcTargetMenuList, , True
        AddProperty "Action: Double Click", "Target Frame", ctcTargetFrameList, , True
        AddProperty "Action: Double Click", "Submenu Alignment", ctcGroupAlignmentList, , True
        AddProperty "Action: Double Click", "New Window Params", ctcNewWindowParams, , True
    End With
    
    IsUpdating = False

End Sub

Private Sub AddProperty(RelativeTo As String, PropertyName As String, CellType As CellTypesConstants, Optional DisableNormal As Boolean, Optional DisableOver As Boolean)

    Dim r As Long
    Dim rp As Long
    Dim npn As String
    Dim WasAdded As Boolean
    
    npn = Space(4) + PropertyName
    
    With mfgData
        .Col = 0
        For r = 1 To .rows - 1
            .Row = r
            If .CellFontBold Then
                If .Text = RelativeTo Then
                    For rp = r + 1 To .rows - 1
                        .Row = rp
                        If .CellFontBold Then
                            Exit For
                        'ElseIf npn < .Text Then
                        '    .AddItem npn
                        '    WasAdded = True
                        '    Exit For
                        End If
                    Next rp
                    Exit For
                End If
            End If
        Next r
        
        If Not WasAdded Then
            .AddItem npn
        End If
        
        .Row = .Row + 1
        .RowData(.Row) = CellType
        .RowHeight(.Row) = 315
        
        If DisableNormal Then
            .Col = 1
            .CellBackColor = Property_DisColor
        End If
        If DisableOver Then
            .Col = 3
            .CellBackColor = Property_DisColor
        End If
        If DisableNormal Or DisableOver Then
            .Col = 2
            .CellBackColor = Property_DisColor
        End If
        
        .Col = 1: .ColSel = 3
        If CellType = ctcNumber Then
            .CellAlignment = flexAlignRightCenter
        Else
            .CellAlignment = flexAlignLeftCenter
        End If
    End With

End Sub

Private Sub AddSection(SectionName As String)

    Dim c As Long
    Dim nsn As String
    
    nsn = Space(0) + SectionName
    
    With mfgData
        .AddItem nsn
        .Row = .rows - 1: .Col = 0
        
        .CellForeColor = vbWhite
        .CellBackColor = Section_DisColor
        .CellAlignment = flexAlignLeftCenter
        '.RowHeight(.Row) = 270
        .CellFontBold = True
        '.CellPictureAlignment = flexAlignLeftCenter
        'Set .CellPicture = ilIcons.ListImages(SectionName).Picture
        
        .Col = 1
        .RowSel = .Row: .ColSel = 3
        .CellBackColor = Section_DisColor
    End With
    
End Sub

Private Sub mfgData_EnterCell()

    If IsUpdating Then Exit Sub
    
    Dim i As Integer

    With mfgData
        If .CellBackColor = Property_DisColor Or .CellBackColor = Section_DisColor Or .Col = 2 Then Exit Sub
        
        Select Case .RowData(.Row)
            Case ctcColor
                Set sCtrl = cmdColor
                sCtrl.BackColor = Hex2Dec(Mid(.Text, 2))
            Case ctcNewWindowParams
                Set sCtrl = cmdWinParams
            Case ctcText
                Set sCtrl = txtText
                sCtrl.Alignment = vbLeftJustify
                sCtrl.Text = .Text
            Case ctcYesNo
                Set sCtrl = cmbCombo
                sCtrl.Clear
                sCtrl.AddItem "Yes"
                sCtrl.AddItem "No"
            Case ctcFont
                Set sCtrl = cmdFont
            Case ctcNumber
                Set sCtrl = txtText
                sCtrl.Alignment = vbRightJustify
                sCtrl.Text = .Text
            Case ctcAlignmentList
                Set sCtrl = cmbCombo
                sCtrl.Clear
                sCtrl.AddItem "Left"
                sCtrl.AddItem "Right"
                sCtrl.AddItem "Center"
            Case ctcGroupAlignmentList
                Set sCtrl = cmbCombo
                sCtrl.Clear
                sCtrl.AddItem "Bottom / Left"
                sCtrl.AddItem "Bottom / Right"
                sCtrl.AddItem "Top / Left"
                sCtrl.AddItem "Top / Right"
                sCtrl.AddItem "Left / Top"
                sCtrl.AddItem "Left / Bottom"
                sCtrl.AddItem "Right / Top"
                sCtrl.AddItem "Right / Bottom"
            Case ctcFile
                Set sCtrl = cmdBrowse
            Case ctcCursorList
                Set sCtrl = cmbCombo
                sCtrl.Clear
                sCtrl.AddItem "Default"
                sCtrl.AddItem "Crosshair"
                sCtrl.AddItem "Hand"
                sCtrl.AddItem "Text"
                sCtrl.AddItem "Help"
            Case ctcActionTypeList
                Set sCtrl = cmbCombo
                sCtrl.Clear
                sCtrl.AddItem "(none)"
                sCtrl.AddItem "Open URL / Execute Script"
                sCtrl.AddItem "Display SubMenu"
                sCtrl.AddItem "Open URL in new Window"
            Case ctcTargetMenuList
                Set sCtrl = cmbCombo
                sCtrl.Clear
                For i = 1 To UBound(MenuGrps)
                    sCtrl.AddItem MenuGrps(i).Name
                Next i
            Case ctcTargetFrameList
                Set sCtrl = cmbCombo
                sCtrl.Clear
                For i = 1 To UBound(FramesInfo.Frames)
                    sCtrl.AddItem FramesInfo.Frames(i)
                Next i
                sCtrl.AddItem "_self"
                sCtrl.AddItem "_top"
        End Select
        
        If TypeOf sCtrl Is ComboBox Then
            For i = 0 To sCtrl.ListCount - 1
                If sCtrl.List(i) = .Text Then
                    sCtrl.ListIndex = i
                    Exit For
                End If
            Next i
        End If
        
        sCtrl.Visible = True
        If TypeOf sCtrl Is ComboBox Or TypeOf sCtrl Is TextBox Then
            sCtrl.Move .CellLeft, .CellTop, .CellWidth
            sCtrl.SetFocus
        Else
            sCtrl.Move .CellLeft + .CellWidth - sCtrl.Width, .CellTop
            Me.SetFocus
        End If
    End With

End Sub

Private Sub mfgData_LeaveCell()

    If IsUpdating Or sCtrl Is Nothing Then Exit Sub

    sCtrl.Visible = False
    
    If TypeOf sCtrl Is TextBox Or TypeOf sCtrl Is ComboBox Then
        mfgData.Text = sCtrl.Text
    End If
    
    Set sCtrl = Nothing

End Sub

Private Sub mfgData_Scroll()

    If IsUpdating Or sCtrl Is Nothing Then Exit Sub
    
    With mfgData
        If TypeOf sCtrl Is ComboBox Or TypeOf sCtrl Is TextBox Then
            sCtrl.Move .CellLeft, .CellTop, .CellWidth
        Else
            sCtrl.Move .CellLeft + .CellWidth - sCtrl.Width, .CellTop, .CellHeight + 30, .CellHeight
        End If
    End With
    
End Sub
