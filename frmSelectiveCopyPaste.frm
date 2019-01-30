VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{9A2947B4-87EE-40EE-A3EF-32BDC32D5726}#1.0#0"; "xfxline3d.ocx"
Begin VB.Form frmSelectiveCopyPaste 
   Caption         =   "Selective Copy/Paste"
   ClientHeight    =   6030
   ClientLeft      =   8820
   ClientTop       =   4305
   ClientWidth     =   5490
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
   ScaleHeight     =   6030
   ScaleWidth      =   5490
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4140
      TabIndex        =   5
      Top             =   5415
      Width           =   885
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
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
      Height          =   345
      Left            =   3075
      TabIndex        =   4
      Top             =   5415
      Width           =   885
   End
   Begin MSComctlLib.StatusBar sbDummy 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   5775
      Width           =   5490
      _ExtentX        =   9684
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
   End
   Begin xfxLine3D.ucLine3D uc3DLineSep 
      Height          =   30
      Left            =   75
      TabIndex        =   2
      Top             =   3330
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   53
   End
   Begin MSComctlLib.TreeView tvProperties 
      Height          =   2775
      Left            =   690
      TabIndex        =   0
      Top             =   60
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   4895
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Checkboxes      =   -1  'True
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
   Begin VB.Frame frmPasteOptions 
      Caption         =   "Paste Options"
      Height          =   1500
      Left            =   75
      TabIndex        =   3
      Top             =   3435
      Width           =   4965
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   90
         ScaleHeight     =   1215
         ScaleWidth      =   4785
         TabIndex        =   7
         Top             =   210
         Width           =   4785
         Begin VB.CommandButton cmdAdvanced 
            Caption         =   "Advanced..."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   270
            TabIndex        =   11
            Top             =   885
            Width           =   1050
         End
         Begin VB.TextBox txtOp 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
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
            Index           =   0
            Left            =   270
            Locked          =   -1  'True
            MousePointer    =   1  'Arrow
            TabIndex        =   10
            Text            =   "Paste to selected"
            Top             =   90
            Width           =   4515
         End
         Begin VB.TextBox txtOp 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
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
            Index           =   1
            Left            =   270
            Locked          =   -1  'True
            MousePointer    =   1  'Arrow
            TabIndex        =   9
            Text            =   "Paste to all under..."
            Top             =   360
            Width           =   4515
         End
         Begin VB.TextBox txtOp 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
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
            Index           =   2
            Left            =   270
            Locked          =   -1  'True
            MousePointer    =   1  'Arrow
            TabIndex        =   8
            Text            =   "Paste to all commands/groups"
            Top             =   615
            Width           =   4515
         End
         Begin VB.OptionButton opPaste 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   0
            TabIndex        =   15
            Top             =   105
            Value           =   -1  'True
            WhatsThisHelpID =   20450
            Width           =   255
         End
         Begin VB.OptionButton opPaste 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   0
            TabIndex        =   14
            Top             =   375
            WhatsThisHelpID =   20460
            Width           =   255
         End
         Begin VB.OptionButton opPaste 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   2
            Left            =   0
            TabIndex        =   13
            Top             =   630
            WhatsThisHelpID =   20470
            Width           =   255
         End
         Begin VB.OptionButton opPaste 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   3
            Left            =   0
            TabIndex        =   12
            Top             =   900
            WhatsThisHelpID =   20470
            Width           =   255
         End
      End
   End
   Begin VB.Label lblMsg 
      BackStyle       =   0  'Transparent
      Caption         =   "Select the properties to copy"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   690
      TabIndex        =   1
      Top             =   2880
      Width           =   4350
   End
   Begin VB.Image imgPaste 
      Height          =   480
      Left            =   135
      Picture         =   "frmSelectiveCopyPaste.frx":0000
      Top             =   885
      Width           =   480
   End
   Begin VB.Image imgCopy 
      Height          =   480
      Left            =   135
      Picture         =   "frmSelectiveCopyPaste.frx":08CA
      Top             =   885
      Width           =   480
   End
   Begin VB.Menu mnuOp 
      Caption         =   "mnuOp"
      Begin VB.Menu mnuOpSel 
         Caption         =   "Select"
         Begin VB.Menu mnuOpSelAll 
            Caption         =   "All"
         End
         Begin VB.Menu mnuOpSelNone 
            Caption         =   "None"
         End
      End
      Begin VB.Menu mnuOpTree 
         Caption         =   "Tree"
         Begin VB.Menu mnuOpTreeExpandThis 
            Caption         =   "Expand"
         End
         Begin VB.Menu mnuOpTreeCollapseThis 
            Caption         =   "Collapse"
         End
         Begin VB.Menu mnuOpTreeSep01 
            Caption         =   "-"
         End
         Begin VB.Menu mnuOpTreeExpand 
            Caption         =   "Expand All"
         End
         Begin VB.Menu mnuOpTreeCollapse 
            Caption         =   "Collapse All"
         End
      End
      Begin VB.Menu mnuOpSep01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpDesc 
         Caption         =   "&Description..."
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmSelectiveCopyPaste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function IsChecked(NodeKey As String) As Boolean

    On Error Resume Next
    IsChecked = tvProperties.Nodes(NodeKey).Checked

End Function

Friend Sub RemoveUnselectedNodes()

    Dim i As Integer
    
    With tvProperties.Nodes
        For i = .Count To 1 Step -1
            If Not .item(i).Checked Then .Remove i
        Next i
    End With

End Sub

Friend Sub CreateNodes()

    Dim nNode As Node
    
    With tvProperties.Nodes
        .Clear
        Set nNode = .Add(, , "ALL", GetLocalizedStr(507), IconIndex("All Properties")):                                                                             nNode.tag = "Select All Properties of %IN%": nNode.Expanded = True
            Set nNode = .Add("ALL", tvwChild, "Color", GetLocalizedStr(212), IconIndex("Color")):                                                                   nNode.tag = "Color Properties of %IN%"
                Set nNode = .Add("Color", tvwChild, "ColorNormal", GetLocalizedStr(179), IconIndex("Normal")):                                                      nNode.tag = "Color Properties of %IN% when the mouse is not over it"
                    Set nNode = .Add("ColorNormal", tvwChild, "ColorNormalText", GetLocalizedStr(196), IconIndex("Text Color")):                                    nNode.tag = "Text Color of %IN% when the mouse is not over it"
                    Set nNode = .Add("ColorNormal", tvwChild, "ColorNormalBack", GetLocalizedStr(508), IconIndex("Back Color")):                                    nNode.tag = "Background Color of %IN% when the mouse is not over it"
                Set nNode = .Add("Color", tvwChild, "ColorOver", GetLocalizedStr(105), IconIndex("Over")):                                                          nNode.tag = "Color Properties of %IN% when the mouse is over it"
                    Set nNode = .Add("ColorOver", tvwChild, "ColorOverText", GetLocalizedStr(196), IconIndex("Text Color")):                                        nNode.tag = "Text Color of %IN% when the mouse is not over it"
                    Set nNode = .Add("ColorOver", tvwChild, "ColorOverBack", GetLocalizedStr(508), IconIndex("Back Color")):                                        nNode.tag = "Background Color of %IN% when the mouse is not over it"
                Set nNode = .Add("Color", tvwChild, "ColorBorder", GetLocalizedStr(355), IconIndex("Frame")):                                                       nNode.tag = "Border size and colors of the %IN% group"
                    Set nNode = .Add("ColorBorder", tvwChild, "ColorBorderSize", GetLocalizedStr(203), IconIndex("Size")):                                          nNode.tag = "Border size of the %IN% group"
                    Set nNode = .Add("ColorBorder", tvwChild, "ColorBorderColors", GetLocalizedStr(207), IconIndex("Color")):                                       nNode.tag = "Color of the borders of the %IN% group"
                    Set nNode = .Add("ColorBorder", tvwChild, "ColorBorderStyle", GetLocalizedStr(613), IconIndex("Highlight Effects")):                            nNode.tag = "Style used for the border in the %IN% group"
                    Set nNode = .Add("ColorBorder", tvwChild, "ColorBorderRadius", "Radius", IconIndex("Margins")):                                                 nNode.tag = "Border Radius used in the %IN% group"
                Set nNode = .Add("Color", tvwChild, "ColorLine", GetLocalizedStr(509), IconIndex("Line")):                                                          nNode.tag = "Foreground Color of the %IN% line"
                Set nNode = .Add("Color", tvwChild, "ColorBack", GetLocalizedStr(508), IconIndex("Back Color")):                                                    nNode.tag = "Background Color of the %IN% group"
                Set nNode = .Add("Color", tvwChild, "ColorToolbarItem", GetLocalizedStr(204), IconIndex("Toolbar Item")):                                           nNode.tag = "Color Properties of the %IN% group when used in the Toolbar"
                    Set nNode = .Add("ColorToolbarItem", tvwChild, "ColorToolbarItemNormal", GetLocalizedStr(179), IconIndex("Normal")):                            nNode.tag = "Color Properties of the %IN% group when used in the Toolbar and the mouse is not over it"
                        Set nNode = .Add("ColorToolbarItemNormal", tvwChild, "ColorToolbarItemNormalText", GetLocalizedStr(196), IconIndex("Text Color")):          nNode.tag = "Text Color of the %IN% group when used in the Toolbar and the mouse is not over it"
                        Set nNode = .Add("ColorToolbarItemNormal", tvwChild, "ColorToolbarItemNormalBack", GetLocalizedStr(508), IconIndex("Back Color")):          nNode.tag = "Background Color of the %IN% group when used in the Toolbar and the mouse is not over it"
                    Set nNode = .Add("ColorToolbarItem", tvwChild, "ColorToolbarItemOver", GetLocalizedStr(105), IconIndex("Over")):                                nNode.tag = "Color Properties of the %IN% group when used in the Toolbar and the mouse is over it"
                        Set nNode = .Add("ColorToolbarItemOver", tvwChild, "ColorToolbarItemOverText", GetLocalizedStr(196), IconIndex("Text Color")):              nNode.tag = "Text Color of the %IN% group when used in the Toolbar and the mouse is over it"
                        Set nNode = .Add("ColorToolbarItemOver", tvwChild, "ColorToolbarItemOverBack", GetLocalizedStr(508), IconIndex("Back Color")):              nNode.tag = "Background Color of the %IN% group when used in the Toolbar and the mouse is over it"
            Set nNode = .Add("ALL", tvwChild, "SeparatorLength", GetLocalizedStr(791), IconIndex("SeparatorLength")):                                               nNode.tag = "Color Properties of %IN%"
            Set nNode = .Add("ALL", tvwChild, "Font", GetLocalizedStr(213), IconIndex("Font")):                                                                     nNode.tag = "Font Properties of %IN%"
                Set nNode = .Add("Font", tvwChild, "FontNormal", GetLocalizedStr(179), IconIndex("Normal")):                                                        nNode.tag = "Font Properties of %IN% when the mouse is not over it"
                    Set nNode = .Add("FontNormal", tvwChild, "nFontName", GetLocalizedStr(409), IconIndex("Font Name")):                                            nNode.tag = "Font Name used by %IN% when the mouse is not over it"
                    Set nNode = .Add("FontNormal", tvwChild, "nFontBold", GetLocalizedStr(510), IconIndex("Font Bold")):                                            nNode.tag = "Bold setting used by %IN% when the mouse is not over it"
                    Set nNode = .Add("FontNormal", tvwChild, "nFontItalic", GetLocalizedStr(511), IconIndex("Font Italic")):                                        nNode.tag = "Italic setting used by %IN% when the mouse is not over it"
                    Set nNode = .Add("FontNormal", tvwChild, "nFontUnderline", GetLocalizedStr(512), IconIndex("Font Underline")):                                  nNode.tag = "Underline setting used by %IN% when the mouse is not over it"
                    Set nNode = .Add("FontNormal", tvwChild, "nFontSize", GetLocalizedStr(203), IconIndex("Size")):                                                 nNode.tag = "Font Size used by %IN% when the mouse is not over it"
                    Set nNode = .Add("FontNormal", tvwChild, "nFontShadow", "Text Shadow", IconIndex("Shadow")):                                                    nNode.tag = "Shadow used by %IN% when the mouse is not over it"
                Set nNode = .Add("Font", tvwChild, "FontOver", GetLocalizedStr(105), IconIndex("Over")):                                                            nNode.tag = "Font Properties of %IN% when the mouse is over it"
                    Set nNode = .Add("FontOver", tvwChild, "oFontName", GetLocalizedStr(409), IconIndex("Font Name")):                                              nNode.tag = "Font Name used by %IN% when the mouse is over it"
                    Set nNode = .Add("FontOver", tvwChild, "oFontBold", GetLocalizedStr(510), IconIndex("Font Bold")):                                              nNode.tag = "Bold setting used by %IN% when the mouse is over it"
                    Set nNode = .Add("FontOver", tvwChild, "oFontItalic", GetLocalizedStr(511), IconIndex("Font Italic")):                                          nNode.tag = "Italic setting used by %IN% when the mouse is over it"
                    Set nNode = .Add("FontOver", tvwChild, "oFontUnderline", GetLocalizedStr(512), IconIndex("Font Underline")):                                    nNode.tag = "Underline setting used by %IN% when the mouse is over it"
                    Set nNode = .Add("FontOver", tvwChild, "oFontSize", GetLocalizedStr(203), IconIndex("Size")):                                                   nNode.tag = "Font Size used by %IN% when the mouse is over it"
                    Set nNode = .Add("FontOver", tvwChild, "oFontShadow", "Text Shadow", IconIndex("Shadow")):                                                      nNode.tag = "Shadow used by %IN% when the mouse is over it"
                Set nNode = .Add("Font", tvwChild, "Alignment", GetLocalizedStr(115), IconIndex("Caption Alignment")):                                              nNode.tag = "Text Alignment used by %IN%"
            Set nNode = .Add("ALL", tvwChild, "Cursor", GetLocalizedStr(215), IconIndex("Cursor")):                                                                 nNode.tag = "Cursor used by %IN% when the mouse is over it"
            Set nNode = .Add("ALL", tvwChild, "Image", GetLocalizedStr(214), IconIndex("Image")):                                                                   nNode.tag = "Image Properties of %IN%"
                Set nNode = .Add("Image", tvwChild, "ImageLeft", GetLocalizedStr(190), IconIndex("Left")):                                                          nNode.tag = "Left Image Properties of %IN%"
                    Set nNode = .Add("ImageLeft", tvwChild, "ImageLeftNormal", GetLocalizedStr(179), IconIndex("Normal")):                                          nNode.tag = "Left Image used by %IN% when the mouse is not over it"
                    Set nNode = .Add("ImageLeft", tvwChild, "ImageLeftOver", GetLocalizedStr(105), IconIndex("Over")):                                              nNode.tag = "Left Image used by %IN% when the mouse is over it"
                    Set nNode = .Add("ImageLeft", tvwChild, "ImageLeftSize", GetLocalizedStr(203), IconIndex("Size")):                                              nNode.tag = "Left Image dimensions used by %IN%"
                    Set nNode = .Add("ImageLeft", tvwChild, "ImageLeftMargin", GetLocalizedStr(988), IconIndex("Margins")):                                         nNode.tag = "Left Image margin used by %IN%"
                Set nNode = .Add("Image", tvwChild, "ImageRight", GetLocalizedStr(191), IconIndex("Right")):                                                        nNode.tag = "Right Image Properties of %IN%"
                    Set nNode = .Add("ImageRight", tvwChild, "ImageRightNormal", GetLocalizedStr(179), IconIndex("Normal")):                                        nNode.tag = "Right Image used by %IN% when the mouse is not over it"
                    Set nNode = .Add("ImageRight", tvwChild, "ImageRightOver", GetLocalizedStr(105), IconIndex("Over")):                                            nNode.tag = "Right Image used by %IN% when the mouse is over it"
                    Set nNode = .Add("ImageRight", tvwChild, "ImageRightSize", GetLocalizedStr(203), IconIndex("Size")):                                            nNode.tag = "Right Image dimensions used by %IN%"
                    Set nNode = .Add("ImageRight", tvwChild, "ImageRightMargin", GetLocalizedStr(988), IconIndex("Margins")):                                       nNode.tag = "Right Image margin used by %IN%"
                Set nNode = .Add("Image", tvwChild, "ImageBorder", GetLocalizedStr(355), IconIndex("Frame")):                                                       nNode.tag = "Border size and colors of the %IN% group"
                    Set nNode = .Add("ImageBorder", tvwChild, "ImageBorderSize", GetLocalizedStr(203), IconIndex("Size")):                                          nNode.tag = "Border size of the %IN% group"
                    Set nNode = .Add("ImageBorder", tvwChild, "ImageBorderColors", GetLocalizedStr(207), IconIndex("Color")):                                       nNode.tag = "Color of the borders of the %IN% group"
                    Set nNode = .Add("ImageBorder", tvwChild, "ImageBorderImages", GetLocalizedStr(672), IconIndex("Image")):                                       nNode.tag = "Images of the borders of the %IN% group"
                        'Set nNode = .Add("ImageBorderImages", tvwChild, "ImageBorderImagesOverlay", GetLocalizedStr(675), IconIndex("Overlay")):                   nNode.Tag = "Overlay amount applied to the Images of the borders of the %IN% group"
                Set nNode = .Add("Image", tvwChild, "ImageBack", GetLocalizedStr(513), IconIndex("Back Color")):                                                    nNode.tag = "Background Image(s) of the %IN% item"
                    Set nNode = .Add("ImageBack", tvwChild, "ImageBackNormal", GetLocalizedStr(179), IconIndex("Normal")):                                          nNode.tag = "Background Image of the %IN% command when the mouse is not over it"
                    Set nNode = .Add("ImageBack", tvwChild, "ImageBackOver", GetLocalizedStr(105), IconIndex("Over")):                                              nNode.tag = "Background Image of the %IN% command when the mouse is over it"
                Set nNode = .Add("Image", tvwChild, "ImageToolbarItem", GetLocalizedStr(204), IconIndex("Toolbar Item")):                                           nNode.tag = "Image Properties of the %IN% group when used in the Toolbar"
                    Set nNode = .Add("ImageToolbarItem", tvwChild, "ImageToolbarItemLeft", GetLocalizedStr(190), IconIndex("Left")):                                nNode.tag = "Left Image Properties of the %IN% group when used in the Toolbar"
                        Set nNode = .Add("ImageToolbarItemLeft", tvwChild, "ImageToolbarItemLeftNormal", GetLocalizedStr(179), IconIndex("Normal")):                nNode.tag = "Left Image used by the %IN% group when used in the Toolbar and the mouse is not over it"
                        Set nNode = .Add("ImageToolbarItemLeft", tvwChild, "ImageToolbarItemLeftOver", GetLocalizedStr(105), IconIndex("Over")):                    nNode.tag = "Left Image used by the %IN% group when used in the Toolbar and the mouse is over it"
                        Set nNode = .Add("ImageToolbarItemLeft", tvwChild, "ImageToolbarItemLeftSize", GetLocalizedStr(203), IconIndex("Size")):                    nNode.tag = "Left Image dimensions used by the %IN% group when used in the Toolbar"
                        Set nNode = .Add("ImageToolbarItemLeft", tvwChild, "ImageToolbarItemLeftMargin", GetLocalizedStr(988), IconIndex("Margins")):               nNode.tag = "Left Image margin used by the %IN% group when used in the Toolbar"
                    Set nNode = .Add("ImageToolbarItem", tvwChild, "ImageToolbarItemRight", GetLocalizedStr(191), IconIndex("Right")):                              nNode.tag = "Right Image Properties of the %IN% group when used in the Toolbar"
                        Set nNode = .Add("ImageToolbarItemRight", tvwChild, "ImageToolbarItemRightNormal", GetLocalizedStr(179), IconIndex("Normal")):              nNode.tag = "Right Image used by the %IN% group when used in the Toolbar and the mouse is not over it"
                        Set nNode = .Add("ImageToolbarItemRight", tvwChild, "ImageToolbarItemRightOver", GetLocalizedStr(105), IconIndex("Over")):                  nNode.tag = "Right Image used by the %IN% group when used in the Toolbar and the mouse is over it"
                        Set nNode = .Add("ImageToolbarItemRight", tvwChild, "ImageToolbarItemRightSize", GetLocalizedStr(203), IconIndex("Size")):                  nNode.tag = "Right Image dimensions used by the %IN% group when used in the Toolbar"
                        Set nNode = .Add("ImageToolbarItemRight", tvwChild, "ImageToolbarItemRightMargin", GetLocalizedStr(988), IconIndex("Margins")):             nNode.tag = "Right Image margin used by the %IN% group when used in the Toolbar"
                    Set nNode = .Add("ImageToolbarItem", tvwChild, "ImageToolbarItemBack", GetLocalizedStr(513), IconIndex("Back Color")):                          nNode.tag = "Background Image Properties of the %IN% group when used in the Toolbar"
                        Set nNode = .Add("ImageToolbarItemBack", tvwChild, "ImageToolbarItemBackNormal", GetLocalizedStr(179), IconIndex("Normal")):                nNode.tag = "Background Image used by the %IN% group when used in the Toolbar and the mouse is not over it"
                        Set nNode = .Add("ImageToolbarItemBack", tvwChild, "ImageToolbarItemBackOver", GetLocalizedStr(105), IconIndex("Over")):                    nNode.tag = "Background Image used by the %IN% group when used in the Toolbar and the mouse is over it"
            Set nNode = .Add("ALL", tvwChild, "Leading", GetLocalizedStr(295), IconIndex("Leading")):                                                               nNode.tag = "Leading used by the %IN% group"
            
            Set nNode = .Add("ALL", tvwChild, "Events", GetLocalizedStr(514), IconIndex("Events")):                                                                 nNode.tag = "Event's Actions of %IN%"
                Set nNode = .Add("Events", tvwChild, "OnClick", GetLocalizedStr(106), IconIndex("EventClick")):                                                     nNode.tag = "Action performed when %IN% is clicked"
                    Set nNode = .Add("OnClick", tvwChild, "OnClickType", GetLocalizedStr(108), IconIndex("Action Type")):                                           nNode.tag = "Type of action performed when %IN% is clicked"
                    Set nNode = .Add("OnClick", tvwChild, "OnClickURL", GetLocalizedStr(111), IconIndex("URL")):                                                    nNode.tag = "URL or script that is executed when %IN% is clicked"
                    Set nNode = .Add("OnClick", tvwChild, "OnClickTargetFrame", GetLocalizedStr(235), IconIndex("Target Frame")):                                   nNode.tag = "Target Frame where the Action should be executed when %IN% is clicked"
                    Set nNode = .Add("OnClick", tvwChild, "OnClickTargetMenu", GetLocalizedStr(109), IconIndex("Group")):                                           nNode.tag = "Menu that must be displayed when %IN% is clicked and the Action Type is set to Display SubMenu"
                    Set nNode = .Add("OnClick", tvwChild, "OnClickNWP", GetLocalizedStr(240), IconIndex("New Window")):                                             nNode.tag = "Parameters of the new window opened when %IN% is clicked and the Action Type is set to Open in new Window"
                    Set nNode = .Add("OnClick", tvwChild, "OnClickTargetMenuAlignment", GetLocalizedStr(725), IconIndex("Group Alignment")):                        nNode.tag = "Alignment of the SubMenu"
                Set nNode = .Add("Events", tvwChild, "OnMouseOver", GetLocalizedStr(180), IconIndex("EventOver")):                                                  nNode.tag = "Action performed when the mouse is moved over %IN%"
                    Set nNode = .Add("OnMouseOver", tvwChild, "OnMouseOverType", GetLocalizedStr(108), IconIndex("Action Type")):                                   nNode.tag = "Type of action performed when the mouse is moved over %IN%"
                    Set nNode = .Add("OnMouseOver", tvwChild, "OnMouseOverURL", GetLocalizedStr(111), IconIndex("URL")):                                            nNode.tag = "URL or script that is executed when the mouse is moved over %IN%"
                    Set nNode = .Add("OnMouseOver", tvwChild, "OnMouseOverTargetFrame", GetLocalizedStr(235), IconIndex("Target Frame")):                           nNode.tag = "Target Frame where the Action should be executed when the mouse is moved over %IN%"
                    Set nNode = .Add("OnMouseOver", tvwChild, "OnMouseOverTargetMenu", GetLocalizedStr(109), IconIndex("Group")):                                   nNode.tag = "Menu that must be displayed when the mouse is moved over %IN% and the Action Type is set to Display SubMenu"
                    Set nNode = .Add("OnMouseOver", tvwChild, "OnMouseOverNWP", GetLocalizedStr(240), IconIndex("New Window")):                                     nNode.tag = "Parameters of the new window opened when the mouse passes over %IN% and the Action Type is set to Open in new Window"
                    Set nNode = .Add("OnMouseOver", tvwChild, "OnMouseOverTargetMenuAlignment", GetLocalizedStr(725), IconIndex("Group Alignment")):                nNode.tag = "Alignment of the SubMenu"
                Set nNode = .Add("Events", tvwChild, "OnDoubleClick", GetLocalizedStr(107), IconIndex("EventDoubleClick")):                                         nNode.tag = "Action performed when %IN% is double clicked"
                    Set nNode = .Add("OnDoubleClick", tvwChild, "OnDoubleClickType", GetLocalizedStr(108), IconIndex("Action Type")):                               nNode.tag = "Type of action performed when %IN% is double clicked"
                    Set nNode = .Add("OnDoubleClick", tvwChild, "OnDoubleClickURL", GetLocalizedStr(111), IconIndex("URL")):                                        nNode.tag = "URL or script that is executed when %IN% is double clicked"
                    Set nNode = .Add("OnDoubleClick", tvwChild, "OnDoubleClickTargetFrame", GetLocalizedStr(235), IconIndex("Target Frame")):                       nNode.tag = "Target Frame where the Action should be executed when %IN% is double clicked"
                    Set nNode = .Add("OnDoubleClick", tvwChild, "OnDoubleClickTargetMenu", GetLocalizedStr(109), IconIndex("Group")):                               nNode.tag = "Menu that must be displayed when %IN% is double clicked and the Action Type is set to Display SubMenu"
                    Set nNode = .Add("OnDoubleClick", tvwChild, "OnDoubleClickNWP", GetLocalizedStr(240), IconIndex("New Window")):                                 nNode.tag = "Parameters of the new window opened when %IN% is double clicked and the Action Type is set to Open in new Window"
                    Set nNode = .Add("OnDoubleClick", tvwChild, "OnDoubleClickTargetMenuAlignment", GetLocalizedStr(725), IconIndex("Group Alignment")):            nNode.tag = "Alignment of the SubMenu"
                
            Set nNode = .Add("ALL", tvwChild, "MenuAlignment", GetLocalizedStr(515), IconIndex("Group Alignment")):                                                 nNode.tag = "Alignment mode used to display the %IN% group"
            Set nNode = .Add("ALL", tvwChild, "Margins", GetLocalizedStr(216), IconIndex("Margins")):                                                               nNode.tag = "Special Effects Properties of the %IN% group"
                Set nNode = .Add("Margins", tvwChild, "MarginsH", GetLocalizedStr(211), IconIndex("Horizontal Margin")):                                            nNode.tag = "Margins of the contents of the %IN% group"
                Set nNode = .Add("Margins", tvwChild, "MarginsV", GetLocalizedStr(210), IconIndex("Vertical Margin")):                                              nNode.tag = "Horizontal margin of the contents of the %IN% group"
            
            Set nNode = .Add("ALL", tvwChild, "SFXCommandHE", GetLocalizedStr(984), IconIndex("Highlight Effects")):                                                nNode.tag = "Selection Effect of the %IN% item"
                    Set nNode = .Add("SFXCommandHE", tvwChild, "SFXCommandHEUseCBNormal", GetLocalizedStr(179), IconIndex("Normal")):                               nNode.tag = "Border color of the %IN% item when the mouse is not over it"
                    Set nNode = .Add("SFXCommandHE", tvwChild, "SFXCommandHEUseCBOver", GetLocalizedStr(105), IconIndex("Over")):                                   nNode.tag = "Border color of the %IN% item when the mouse is over it"
                    Set nNode = .Add("SFXCommandHE", tvwChild, "SFXCommandHEBorderSize", GetLocalizedStr(206), IconIndex("Border Size")):                           nNode.tag = "Border Size of the %IN% item"
                    Set nNode = .Add("SFXCommandHE", tvwChild, "SFXCommandHEMarginX", GetLocalizedStr(219), IconIndex("Command Horizontal Margin")):                nNode.tag = "Horizontal Margin of the %IN% item"
                    Set nNode = .Add("SFXCommandHE", tvwChild, "SFXCommandHEMarginY", GetLocalizedStr(220), IconIndex("Command Vertical Margin")):                  nNode.tag = "Vertical Margin of the %IN% item"
                    Set nNode = .Add("SFXCommandHE", tvwChild, "SFXCommandHERadius", "Radius", IconIndex("Margins")):                  nNode.tag = "Radius of the %IN% item"
            
            #If LITE = 0 Then
            Set nNode = .Add("ALL", tvwChild, "SFX", GetLocalizedStr(231), IconIndex("Special Effects")):                                                           nNode.tag = "Vertical margin of the contents of the %IN% group"
                Set nNode = .Add("SFX", tvwChild, "SFXGroupSize", GetLocalizedStr(227), IconIndex("Size")):                                                         nNode.tag = "Group Width setting of the %IN% group"
                    Set nNode = .Add("SFXGroupSize", tvwChild, "SFXGroupWidth", GetLocalizedStr(516), IconIndex("Group Width")):                                    nNode.tag = "Group Size Parameters for the %IN% group"
                    Set nNode = .Add("SFXGroupSize", tvwChild, "SFXGroupHeight", GetLocalizedStr(517), IconIndex("Group Height")):                                  nNode.tag = "Group Height setting of the %IN% group"
                    Set nNode = .Add("SFXGroupSize", tvwChild, "SFXGroupScrolling", GetLocalizedStr(830), IconIndex("Scrolling")):                                  nNode.tag = "Group Height setting of the %IN% group"
                Set nNode = .Add("SFX", tvwChild, "SFXEx", GetLocalizedStr(218), IconIndex("Group Effects")):                                                       nNode.tag = "Special Group Effects used by the %IN% group"
                    Set nNode = .Add("SFXEx", tvwChild, "SFXDropShadow", GetLocalizedStr(831), IconIndex("Shadow")):                                                nNode.tag = "Drop Shadow settings of the %IN% group"
                        Set nNode = .Add("SFXDropShadow", tvwChild, "SFXDropShadowSize", GetLocalizedStr(212), IconIndex("Size")):                                                nNode.tag = "Drop Shadow Size of the %IN% group"
                        Set nNode = .Add("SFXDropShadow", tvwChild, "SFXDropShadowColor", GetLocalizedStr(203), IconIndex("Color")):                                                nNode.tag = "Drop Shadow Color of the %IN% group"
                    Set nNode = .Add("SFXEx", tvwChild, "SFXTransparency", GetLocalizedStr(222), IconIndex("Transparency")):                                        nNode.tag = "Transparency setting of the %IN% group"
            'Set nNode = .Add("ALL", tvwChild, "Sounds", GetLocalizedStr(518), IconIndex("Sounds")):                                                                 nNode.Tag = "Sounds used by the %IN% item"
            '    Set nNode = .Add("Sounds", tvwChild, "SoundsOMO", GetLocalizedStr(519), IconIndex("EventOver")):                                                    nNode.Tag = "Sound played when the user moves the mouse over the %IN% item"
            '    Set nNode = .Add("Sounds", tvwChild, "SoundsOC", GetLocalizedStr(520), IconIndex("EventClick")):                                                    nNode.Tag = "Sound played when the user clicks the %IN% item"
            #End If
            Set nNode = .Add("ALL", tvwChild, "CommandsLayout", GetLocalizedStr(668), IconIndex("Commands Layout")):                                                nNode.tag = "Commands Layout for the %IN% group"
            Set nNode = .Add("ALL", tvwChild, "StatusText", "Status Text", IconIndex("Status Text")):                                                               nNode.tag = "Status/Tooltip to be displayed when the mouse is over %IN%"
            Set nNode = .Add("ALL", tvwChild, "Compile", "Compile", IconIndex("Compile")):                                                                          nNode.tag = "Controls if the item will be compiled or not"
            
            Set nNode = .Add("ALL", tvwChild, "TBAppearance", "Appearance", IconIndex("Toolbar Item")):                                                             nNode.tag = "Appearance of the toolbar"
                Set nNode = .Add("TBAppearance", tvwChild, "TBToolbarStyle", "Toolbar Style", IconIndex("Commands Layout")):                                        nNode.tag = "Style (Horizontal/Vertical) of the toolbar"
                Set nNode = .Add("TBAppearance", tvwChild, "TBBorder", "Border", IconIndex("Frame")):                                                   nNode.tag = "Border of the toolbar"
                    Set nNode = .Add("TBBorder", tvwChild, "TBBorderSize", "Size", IconIndex("Size")):                                                   nNode.tag = "Border Size"
                    Set nNode = .Add("TBBorder", tvwChild, "TBBorderColor", "Color", IconIndex("Color")):                                                   nNode.tag = "Border Color"
                    Set nNode = .Add("TBBorder", tvwChild, "TBBorderStyle", "Style", IconIndex("Highlight Effects")):                                                   nNode.tag = "Border Style"
                Set nNode = .Add("TBAppearance", tvwChild, "TBMargins", "Margins", IconIndex("Margins")):                                                   nNode.tag = "Margins of the toolbar -- control the distance between the toolbar's borders and the inner contents"
                    Set nNode = .Add("TBMargins", tvwChild, "TBHorizontal", "Horizontal", IconIndex("Horizontal Margin")):                                                   nNode.tag = "Horizontal Margin"
                    Set nNode = .Add("TBMargins", tvwChild, "TBVertical", "Vertical", IconIndex("Horizontal Margin")):                                                   nNode.tag = "Vertical Margin"
                Set nNode = .Add("TBAppearance", tvwChild, "TBSeparation", "Separation", IconIndex("Leading")):                                                   nNode.tag = "Separation between the toolbar items"
                Set nNode = .Add("TBAppearance", tvwChild, "TBJustifyHotSpots", "Justify HotSpots", IconIndex("Justify")):                                                    nNode.tag = "Justify (make equal) the width (for horizontal toolbars) or the height (for vertical toolbars) of the toolbar items"
                Set nNode = .Add("TBAppearance", tvwChild, "TBBackColor", "Back Color", IconIndex("Back Color")):                                                    nNode.tag = "Background color of the toolbar"
                Set nNode = .Add("TBAppearance", tvwChild, "TBBackImage", "Back Image", IconIndex("Image")):                                                    nNode.tag = "Background image of the toolbar"
                Set nNode = .Add("TBAppearance", tvwChild, "TBRadius", "Radius", IconIndex("Margins")):                                                    nNode.tag = "Radius of the toolbar"
            Set nNode = .Add("ALL", tvwChild, "TBPositioning", "Positioning", IconIndex("Toolbar Alignment")):                                                   nNode.tag = "Positioning settings and controls for the toolbar"
                Set nNode = .Add("TBPositioning", tvwChild, "TBAlignment", "Alignment", IconIndex("Toolbar Alignment")):                                                    nNode.tag = "Alignment and positioning method for the toolbar"
                Set nNode = .Add("TBPositioning", tvwChild, "TBSpanning", "Spanning", IconIndex("Spanning")):                                                    nNode.tag = "Control how the toolbar background is rendered"
                Set nNode = .Add("TBPositioning", tvwChild, "TBOffset", "Offset", IconIndex("Margins")):                                                    nNode.tag = "Offset of the toolbar"
                    Set nNode = .Add("TBOffset", tvwChild, "TBOHorizontal", "Horizontal", IconIndex("Horizontal Margin")):                                                    nNode.tag = "Horizontal Offset"
                    Set nNode = .Add("TBOffset", tvwChild, "TBOVertical", "Vertical", IconIndex("Vertical Margin")):                                                    nNode.tag = "Vertical Offset"
            Set nNode = .Add("ALL", tvwChild, "TBEffects", "Effects", IconIndex("Special Effects")):                                                    nNode.tag = "Special Effects of the toolbar"
                Set nNode = .Add("TBEffects", tvwChild, "TBDropShadow", GetLocalizedStr(831), IconIndex("Shadow")):                                                nNode.tag = "Drop Shadow settings of the toolbar"
                        Set nNode = .Add("TBDropShadow", tvwChild, "TBDropShadowSize", GetLocalizedStr(212), IconIndex("Size")):                                                nNode.tag = "Drop Shadow Size of the toolbar"
                        Set nNode = .Add("TBDropShadow", tvwChild, "TBDropShadowColor", GetLocalizedStr(203), IconIndex("Color")):                                                nNode.tag = "Drop Shadow Color of the toolbar"
                    Set nNode = .Add("TBEffects", tvwChild, "TBTransparency", GetLocalizedStr(222), IconIndex("Transparency")):                                        nNode.tag = "Transparency setting of the toolbar"
            Set nNode = .Add("ALL", tvwChild, "TBAdvanced", "Advanced", IconIndex("Group Effects")):                                                   nNode.tag = "Advanced settings of the toolbar"
                Set nNode = .Add("TBAdvanced", tvwChild, "TBFollowSB", "Follow Scrollbars", IconIndex("Follow Scrolling")):                                                    nNode.tag = "Follow browser's scrollbars"
                Set nNode = .Add("TBAdvanced", tvwChild, "TBSize", "Size", IconIndex("Size")):                                                    nNode.tag = "Size of the toolbar"
                    Set nNode = .Add("TBSize", tvwChild, "TBWidth", "Width", IconIndex("Group Width")):                                    nNode.tag = "Width of the toolbar"
                    Set nNode = .Add("TBSize", tvwChild, "TBHeight", "Height", IconIndex("Group Height")):                                  nNode.tag = "Height of the toolbar"
                Set nNode = .Add("TBAdvanced", tvwChild, "TBVisCon", "Visibility Condition", IconIndex("Visibility Condition")):                                                    nNode.tag = "Simple javascript function that controls the toolbar's visibility"
    End With
    
    With tvProperties
        If IsGroup(frmMain.tvMenus.SelectedItem.key) Then
            .Nodes.Remove .Nodes("ColorNormal").Index
            .Nodes.Remove .Nodes("ColorOver").Index
            .Nodes.Remove .Nodes("ColorLine").Index
            .Nodes.Remove .Nodes("ImageLeft").Index
            .Nodes.Remove .Nodes("ImageRight").Index
            .Nodes.Remove .Nodes("ImageBackNormal").Index
            .Nodes.Remove .Nodes("ImageBackOver").Index
            If Not CreateToolbar Then
                .Nodes.Remove .Nodes("ColorToolbarItem").Index
                .Nodes.Remove .Nodes("ImageToolbarItem").Index
            End If
            .Nodes.Remove .Nodes("SeparatorLength").Index
        End If
        If IsCommand(frmMain.tvMenus.SelectedItem.key) Then
            .Nodes.Remove .Nodes("ColorBack").Index
            .Nodes.Remove .Nodes("ColorLine").Index
            .Nodes.Remove .Nodes("ColorToolbarItem").Index
            .Nodes.Remove .Nodes("ColorBorder").Index
            .Nodes.Remove .Nodes("ImageBorder").Index
            .Nodes.Remove .Nodes("ImageToolbarItem").Index
            .Nodes.Remove .Nodes("Leading").Index
            .Nodes.Remove .Nodes("MenuAlignment").Index
            .Nodes.Remove .Nodes("SFX").Index
            .Nodes.Remove .Nodes("Margins").Index
            .Nodes.Remove .Nodes("CommandsLayout").Index
            .Nodes.Remove .Nodes("SeparatorLength").Index
        End If
        If IsSeparator(frmMain.tvMenus.SelectedItem.key) Then
            .Nodes.Remove .Nodes("ColorNormal").Index
            .Nodes.Remove .Nodes("ColorOver").Index
            .Nodes.Remove .Nodes("ColorToolbarItem").Index
            .Nodes.Remove .Nodes("ColorBorder").Index
            .Nodes.Remove .Nodes("Font").Index
            .Nodes.Remove .Nodes("Cursor").Index
            .Nodes.Remove .Nodes("Image").Index
            .Nodes.Remove .Nodes("Leading").Index
            .Nodes.Remove .Nodes("Events").Index
            .Nodes.Remove .Nodes("MenuAlignment").Index
            .Nodes.Remove .Nodes("SFX").Index
            .Nodes.Remove .Nodes("Margins").Index
            .Nodes.Remove .Nodes("CommandsLayout").Index
            .Nodes.Remove .Nodes("StatusText").Index
            .Nodes.Remove .Nodes("SFXCommandHE").Index
        End If
        If IsTBMapSel Then
            .Nodes.Remove .Nodes("Color").Index
            .Nodes.Remove .Nodes("Font").Index
            .Nodes.Remove .Nodes("Cursor").Index
            .Nodes.Remove .Nodes("Image").Index
            .Nodes.Remove .Nodes("Leading").Index
            .Nodes.Remove .Nodes("Events").Index
            .Nodes.Remove .Nodes("MenuAlignment").Index
            .Nodes.Remove .Nodes("SFX").Index
            .Nodes.Remove .Nodes("Margins").Index
            .Nodes.Remove .Nodes("CommandsLayout").Index
            .Nodes.Remove .Nodes("StatusText").Index
            .Nodes.Remove .Nodes("SFXCommandHE").Index
        Else
            .Nodes.Remove .Nodes("TBAppearance").Index
            .Nodes.Remove .Nodes("TBPositioning").Index
            .Nodes.Remove .Nodes("TBEffects").Index
            .Nodes.Remove .Nodes("TBAdvanced").Index
        End If
        
        On Error Resume Next
        'DISABLE ALL SOUND FEATURES
        .Nodes.Remove .Nodes("Sounds").Index
    End With
    
End Sub

Private Sub cmdAdvanced_Click()

    Dim i As Integer

    opPaste(3).Value = True
    opPaste(3).SetFocus
    
    frmSelItems.Show vbModal
    
    On Error Resume Next
    i = UBound(dmbClipboard.CustomSel)
    If Err.number Then opPaste(0).Value = True

End Sub

Private Sub cmdCancel_Click()

    Hide

End Sub

Private Sub cmdOK_Click()

    If imgPaste.Visible Then
        If DoPaste = False Then Exit Sub
        frmMain.SaveState GetLocalizedStr(521)
    End If
    
    Hide

End Sub

Private Function DoPaste() As Boolean

    Dim i As Integer
    Dim j As Integer
    Dim nNode As Node
    Dim ns As Integer
    
    If dmbClipboard.ObjSrc = docToolbar Then
        If opPaste(0).Value Then
            DoPasteToolbar dmbClipboard.TBContents, Project.Toolbars(ToolbarIndexByKey(frmMain.tvMapView.SelectedItem.key))
        Else
            For i = 1 To UBound(Project.Toolbars)
                DoPasteToolbar dmbClipboard.TBContents, Project.Toolbars(i)
            Next i
        End If
    Else
        If (dmbClipboard.ObjSrc = docCommand) Or (dmbClipboard.ObjSrc = docSeparator) Then
            If opPaste(0).Value Then
                DoPasteCommand dmbClipboard.CmdContents, MenuCmds(GetID(frmMain.tvMenus.SelectedItem))
            ElseIf opPaste(1).Value Then
                For i = 1 To UBound(MenuCmds)
                    If MenuCmds(i).parent = GetID(frmMain.tvMenus.SelectedItem.parent) Then
                        If (dmbClipboard.ObjSrc = docCommand) Then
                            If MenuCmds(i).Name <> "[SEP]" Then
                                DoPasteCommand dmbClipboard.CmdContents, MenuCmds(i)
                            End If
                        End If
                        If (dmbClipboard.ObjSrc = docSeparator) Then
                            If MenuCmds(i).Name = "[SEP]" Then
                                DoPasteCommand dmbClipboard.CmdContents, MenuCmds(i)
                            End If
                        End If
                    End If
                Next i
            ElseIf opPaste(2).Value Then
                For i = 1 To UBound(MenuCmds)
                    If (dmbClipboard.ObjSrc = docCommand) Then
                        If MenuCmds(i).Name <> "[SEP]" Then
                            DoPasteCommand dmbClipboard.CmdContents, MenuCmds(i)
                        End If
                    End If
                    If (dmbClipboard.ObjSrc = docSeparator) Then
                        If MenuCmds(i).Name = "[SEP]" Then
                            DoPasteCommand dmbClipboard.CmdContents, MenuCmds(i)
                        End If
                    End If
                Next i
            ElseIf opPaste(3).Value Then
                On Error Resume Next
                ns = UBound(dmbClipboard.CustomSel)
                If ns = 0 Then
                    cmdAdvanced_Click
                    DoPaste = False
                    Exit Function
                End If
                
                For i = 1 To UBound(MenuCmds)
                    For j = 1 To ns
                        If MenuCmds(i).Name = dmbClipboard.CustomSel(j) Then
                            DoPasteCommand dmbClipboard.CmdContents, MenuCmds(i)
                        End If
                    Next j
                Next i
            End If
        Else
            If opPaste(0).Value Then
                DoPasteGroup dmbClipboard.GrpContents, MenuGrps(GetID(frmMain.tvMenus.SelectedItem))
            ElseIf opPaste(2).Value Then
                For i = 1 To UBound(MenuGrps)
                    DoPasteGroup dmbClipboard.GrpContents, MenuGrps(i)
                Next i
            ElseIf opPaste(3).Value Then
                On Error Resume Next
                ns = UBound(dmbClipboard.CustomSel)
                If ns = 0 Then
                    cmdAdvanced_Click
                    DoPaste = False
                    Exit Function
                End If
                
                For i = 1 To UBound(MenuGrps)
                    For j = 1 To ns
                        If MenuGrps(i).Name = dmbClipboard.CustomSel(j) Then
                            DoPasteGroup dmbClipboard.GrpContents, MenuGrps(i)
                        End If
                    Next j
                Next i
            End If
        End If
    End If
    
    For Each nNode In frmMain.tvMenus.Nodes
        If IsGroup(nNode.key) Then
            nNode.Image = GenGrpIcon(GetID(nNode))
        Else
            nNode.Image = GenCmdIcon(GetID(nNode))
        End If
    Next nNode
    
    DoPaste = True

End Function

Private Sub DoPasteCommand(Src As MenuCmd, trg As MenuCmd)

    If IsChecked("ColorNormalText") Then trg.nTextColor = Src.nTextColor
    If IsChecked("ColorNormalBack") Then trg.nBackColor = Src.nBackColor
    
    If IsChecked("ColorOverText") Then trg.hTextColor = Src.hTextColor
    If IsChecked("ColorOverBack") Then trg.hBackColor = Src.hBackColor
    
    If IsChecked("ColorLine") Then trg.nTextColor = Src.nTextColor
    If IsChecked("ColorBack") Then trg.nBackColor = Src.nBackColor
    
    If IsChecked("nFontName") Then trg.NormalFont.FontName = Src.NormalFont.FontName
    If IsChecked("nFontBold") Then trg.NormalFont.FontBold = Src.NormalFont.FontBold
    If IsChecked("nFontItalic") Then trg.NormalFont.FontItalic = Src.NormalFont.FontItalic
    If IsChecked("nFontUnderline") Then trg.NormalFont.FontUnderline = Src.NormalFont.FontUnderline
    If IsChecked("nFontSize") Then trg.NormalFont.FontSize = Src.NormalFont.FontSize
    If IsChecked("nFontShadow") Then trg.NormalFont.FontShadow = Src.NormalFont.FontShadow
    
    If IsChecked("oFontName") Then trg.HoverFont.FontName = Src.HoverFont.FontName
    If IsChecked("oFontBold") Then trg.HoverFont.FontBold = Src.HoverFont.FontBold
    If IsChecked("oFontItalic") Then trg.HoverFont.FontItalic = Src.HoverFont.FontItalic
    If IsChecked("oFontUnderline") Then trg.HoverFont.FontUnderline = Src.HoverFont.FontUnderline
    If IsChecked("oFontSize") Then trg.HoverFont.FontSize = Src.HoverFont.FontSize
    If IsChecked("oFontShadow") Then trg.HoverFont.FontShadow = Src.HoverFont.FontShadow
    
    If IsChecked("Alignment") Then trg.Alignment = Src.Alignment
    
    If IsChecked("Cursor") Then trg.iCursor = Src.iCursor
    
    If IsChecked("ImageLeftNormal") Then trg.LeftImage.NormalImage = Src.LeftImage.NormalImage
    If IsChecked("ImageLeftOver") Then trg.LeftImage.HoverImage = Src.LeftImage.HoverImage
    If IsChecked("ImageLeftSize") Then
        trg.LeftImage.w = Src.LeftImage.w
        trg.LeftImage.h = Src.LeftImage.h
    End If
    If IsChecked("ImageLeftMargin") Then trg.LeftImage.margin = Src.LeftImage.margin
    
    If IsChecked("ImageRightNormal") Then trg.RightImage.NormalImage = Src.RightImage.NormalImage
    If IsChecked("ImageRightOver") Then trg.RightImage.HoverImage = Src.RightImage.HoverImage
    If IsChecked("ImageRightSize") Then
        trg.RightImage.w = Src.RightImage.w
        trg.RightImage.h = Src.RightImage.h
    End If
    If IsChecked("ImageRightMargin") Then trg.RightImage.margin = Src.RightImage.margin
    
    If IsChecked("ImageBack") Then
        trg.BackImage.Tile = Src.BackImage.Tile
        trg.BackImage.AllowCrop = Src.BackImage.AllowCrop
    End If
    If IsChecked("ImageBackNormal") Then trg.BackImage.NormalImage = Src.BackImage.NormalImage
    If IsChecked("ImageBackOver") Then trg.BackImage.HoverImage = Src.BackImage.HoverImage
    
    If IsChecked("OnMouseOverType") Then trg.Actions.onmouseover.Type = Src.Actions.onmouseover.Type
    If IsChecked("OnMouseOverURL") Then trg.Actions.onmouseover.url = Src.Actions.onmouseover.url
    If IsChecked("OnMouseOverTargetFrame") Then trg.Actions.onmouseover.TargetFrame = Src.Actions.onmouseover.TargetFrame
    If IsChecked("OnMouseOverTargetMenu") Then trg.Actions.onmouseover.TargetMenu = Src.Actions.onmouseover.TargetMenu
    If IsChecked("OnMouseOverNWP") Then trg.Actions.onmouseover.WindowOpenParams = Src.Actions.onmouseover.WindowOpenParams
    If IsChecked("OnMouseOverTargetMenuAlignment") Then trg.Actions.onmouseover.TargetMenuAlignment = Src.Actions.onmouseover.TargetMenuAlignment
    
    If IsChecked("OnClickType") Then trg.Actions.onclick.Type = Src.Actions.onclick.Type
    If IsChecked("OnClickURL") Then trg.Actions.onclick.url = Src.Actions.onclick.url
    If IsChecked("OnClickTargetFrame") Then trg.Actions.onclick.TargetFrame = Src.Actions.onclick.TargetFrame
    If IsChecked("OnClickTargetMenu") Then trg.Actions.onclick.TargetMenu = Src.Actions.onclick.TargetMenu
    If IsChecked("OnClickNWP") Then trg.Actions.onclick.WindowOpenParams = Src.Actions.onclick.WindowOpenParams
    If IsChecked("OnClickTargetMenuAlignment") Then trg.Actions.onclick.TargetMenuAlignment = Src.Actions.onclick.TargetMenuAlignment
    
    If IsChecked("OnDoubleClickType") Then trg.Actions.OnDoubleClick.Type = Src.Actions.OnDoubleClick.Type
    If IsChecked("OnDoubleClickURL") Then trg.Actions.OnDoubleClick.url = Src.Actions.OnDoubleClick.url
    If IsChecked("OnDoubleClickTargetFrame") Then trg.Actions.OnDoubleClick.TargetFrame = Src.Actions.OnDoubleClick.TargetFrame
    If IsChecked("OnDoubleClickTargetMenu") Then trg.Actions.OnDoubleClick.TargetMenu = Src.Actions.OnDoubleClick.TargetMenu
    If IsChecked("OnDoubleClickNWP") Then trg.Actions.OnDoubleClick.WindowOpenParams = Src.Actions.OnDoubleClick.WindowOpenParams
    If IsChecked("OnDoubleClickTargetMenuAlignment") Then trg.Actions.OnDoubleClick.TargetMenuAlignment = Src.Actions.OnDoubleClick.TargetMenuAlignment
    
    If IsChecked("SFXCommandHE") Then
        trg.CmdsFXNormal = Src.CmdsFXNormal
        trg.CmdsFXOver = Src.CmdsFXOver
    End If
    If IsChecked("SFXCommandHEBorderSize") Then trg.CmdsFXSize = Src.CmdsFXSize
    If IsChecked("SFXCommandHEMarginX") Then trg.CmdsMarginX = Src.CmdsMarginX
    If IsChecked("SFXCommandHEMarginY") Then trg.CmdsMarginY = Src.CmdsMarginY
    If IsChecked("SFXCommandHEUseCBNormal") Then trg.CmdsFXnColor = Src.CmdsFXnColor
    If IsChecked("SFXCommandHEUseCBOver") Then trg.CmdsFXhColor = Src.CmdsFXhColor
    If IsChecked("SFXCommandHERadius") Then trg.Radius = Src.Radius
    
    If IsChecked("SoundsOMO") Then trg.Sound.onmouseover = Src.Sound.onmouseover
    If IsChecked("SoundsOC") Then trg.Sound.onclick = Src.Sound.onclick
    
    If IsChecked("SeparatorLength") Then trg.SeparatorPercent = Src.SeparatorPercent
    
    If IsChecked("StatusText") Then trg.WinStatus = Src.WinStatus
    
    If IsChecked("Compile") Then trg.Compile = Src.Compile
    
End Sub

Private Sub DoPasteGroup(Src As MenuGrp, trg As MenuGrp)

    If IsChecked("ColorBack") Then trg.bColor = Src.bColor
    
    If IsChecked("ColorToolbarItemNormalText") Then trg.nTextColor = Src.nTextColor
    If IsChecked("ColorToolbarItemNormalBack") Then trg.nBackColor = Src.nBackColor
    
    If IsChecked("ColorToolbarItemOverText") Then trg.hTextColor = Src.hTextColor
    If IsChecked("ColorToolbarItemOverBack") Then trg.hBackColor = Src.hBackColor
    
    If IsChecked("ColorBorderSize") Then trg.frameBorder = Src.frameBorder
    If IsChecked("ColorBorderColors") Then trg.Corners = Src.Corners
    If IsChecked("ColorBorderStyle") Then trg.BorderStyle = Src.BorderStyle
    If IsChecked("ColorBorderRadius") Then trg.Radius = Src.Radius
    
    If IsChecked("nFontName") Then trg.DefNormalFont.FontName = Src.DefNormalFont.FontName
    If IsChecked("nFontBold") Then trg.DefNormalFont.FontBold = Src.DefNormalFont.FontBold
    If IsChecked("nFontItalic") Then trg.DefNormalFont.FontItalic = Src.DefNormalFont.FontItalic
    If IsChecked("nFontUnderline") Then trg.DefNormalFont.FontUnderline = Src.DefNormalFont.FontUnderline
    If IsChecked("nFontSize") Then trg.DefNormalFont.FontSize = Src.DefNormalFont.FontSize
    If IsChecked("nFontShadow") Then trg.DefNormalFont.FontShadow = Src.DefNormalFont.FontShadow
    
    If IsChecked("oFontName") Then trg.DefHoverFont.FontName = Src.DefHoverFont.FontName
    If IsChecked("oFontBold") Then trg.DefHoverFont.FontBold = Src.DefHoverFont.FontBold
    If IsChecked("oFontItalic") Then trg.DefHoverFont.FontItalic = Src.DefHoverFont.FontItalic
    If IsChecked("oFontUnderline") Then trg.DefHoverFont.FontUnderline = Src.DefHoverFont.FontUnderline
    If IsChecked("oFontSize") Then trg.DefHoverFont.FontSize = Src.DefHoverFont.FontSize
    If IsChecked("oFontShadow") Then trg.DefHoverFont.FontShadow = Src.DefHoverFont.FontShadow
    
    If IsChecked("Alignment") Then trg.CaptionAlignment = Src.CaptionAlignment
    
    If IsChecked("Cursor") Then trg.iCursor = Src.iCursor
    
    If IsChecked("ImageBorderSize") Then trg.frameBorder = Src.frameBorder
    If IsChecked("ImageBorderColors") Then trg.Corners = Src.Corners
    If IsChecked("ImageBorderImages") Then trg.CornersImages = Src.CornersImages
    
    If IsChecked("ImageBack") Then trg.Image = Src.Image
    
    If IsChecked("ImageToolbarItemLeftNormal") Then trg.tbiLeftImage.NormalImage = Src.tbiLeftImage.NormalImage
    If IsChecked("ImageToolbarItemLeftOver") Then trg.tbiLeftImage.HoverImage = Src.tbiLeftImage.HoverImage
    If IsChecked("ImageToolbarItemLeftSize") Then
        trg.tbiLeftImage.w = Src.tbiLeftImage.w
        trg.tbiLeftImage.h = Src.tbiLeftImage.h
    End If
    If IsChecked("ImageToolbarItemLeftMargin") Then trg.tbiLeftImage.margin = Src.tbiLeftImage.margin
    
    If IsChecked("ImageToolbarItemRightNormal") Then trg.tbiRightImage.NormalImage = Src.tbiRightImage.NormalImage
    If IsChecked("ImageToolbarItemRightOver") Then trg.tbiRightImage.HoverImage = Src.tbiRightImage.HoverImage
    If IsChecked("ImageToolbarItemRightSize") Then
        trg.tbiRightImage.w = Src.tbiRightImage.w
        trg.tbiRightImage.h = Src.tbiRightImage.h
    End If
    If IsChecked("ImageToolbarItemRightMargin") Then trg.tbiRightImage.margin = Src.tbiRightImage.margin
    
    If IsChecked("ImageToolbarItemBack") Then
        trg.tbiBackImage.Tile = Src.tbiBackImage.Tile
        trg.tbiBackImage.AllowCrop = Src.tbiBackImage.AllowCrop
    End If
    If IsChecked("ImageToolbarItemBackNormal") Then trg.tbiBackImage.NormalImage = Src.tbiBackImage.NormalImage
    If IsChecked("ImageToolbarItemBackOver") Then trg.tbiBackImage.HoverImage = Src.tbiBackImage.HoverImage
    
    If IsChecked("Leading") Then trg.Leading = Src.Leading
    
    If IsChecked("OnMouseOverType") Then trg.Actions.onmouseover.Type = Src.Actions.onmouseover.Type
    If IsChecked("OnMouseOverURL") Then trg.Actions.onmouseover.url = Src.Actions.onmouseover.url
    If IsChecked("OnMouseOverTargetFrame") Then trg.Actions.onmouseover.TargetFrame = Src.Actions.onmouseover.TargetFrame
    If IsChecked("OnMouseOverTargetMenu") Then trg.Actions.onmouseover.TargetMenu = Src.Actions.onmouseover.TargetMenu
    
    If IsChecked("OnClickType") Then trg.Actions.onclick.Type = Src.Actions.onclick.Type
    If IsChecked("OnClickURL") Then trg.Actions.onclick.url = Src.Actions.onclick.url
    If IsChecked("OnClickTargetFrame") Then trg.Actions.onclick.TargetFrame = Src.Actions.onclick.TargetFrame
    If IsChecked("OnClickTargetMenu") Then trg.Actions.onclick.TargetMenu = Src.Actions.onclick.TargetMenu
    
    If IsChecked("OnDoubleClickType") Then trg.Actions.OnDoubleClick.Type = Src.Actions.OnDoubleClick.Type
    If IsChecked("OnDoubleClickURL") Then trg.Actions.OnDoubleClick.url = Src.Actions.OnDoubleClick.url
    If IsChecked("OnDoubleClickTargetFrame") Then trg.Actions.OnDoubleClick.TargetFrame = Src.Actions.OnDoubleClick.TargetFrame
    If IsChecked("OnDoubleClickTargetMenu") Then trg.Actions.OnDoubleClick.TargetMenu = Src.Actions.OnDoubleClick.TargetMenu
    
    If IsChecked("MenuAlignment") Then trg.Alignment = Src.Alignment
    
    If IsChecked("MarginsH") Then trg.ContentsMarginH = Src.ContentsMarginH
    If IsChecked("MarginsH") Then trg.ContentsMarginV = Src.ContentsMarginV
    
    If IsChecked("SFXCommandHE") Then
        trg.CmdsFXNormal = Src.CmdsFXNormal
        trg.CmdsFXOver = Src.CmdsFXOver
    End If
    If IsChecked("SFXCommandHEBorderSize") Then trg.CmdsFXSize = Src.CmdsFXSize
    If IsChecked("SFXCommandHEMarginX") Then trg.CmdsMarginX = Src.CmdsMarginX
    If IsChecked("SFXCommandHEMarginY") Then trg.CmdsMarginY = Src.CmdsMarginY
    If IsChecked("SFXCommandHEUseCBNormal") Then trg.CmdsFXnColor = Src.CmdsFXnColor
    If IsChecked("SFXCommandHEUseCBOver") Then trg.CmdsFXhColor = Src.CmdsFXhColor
    If IsChecked("SFXCommandHERadius") Then trg.tbiRadius = Src.tbiRadius
    
    If IsChecked("SFXDropShadowColor") Then trg.DropShadowColor = Src.DropShadowColor
    If IsChecked("SFXDropShadowSize") Then trg.DropShadowSize = Src.DropShadowSize
    If IsChecked("SFXTransparency") Then trg.Transparency = Src.Transparency

    If IsChecked("SFXGroupWidth") Then trg.fWidth = Src.fWidth
    If IsChecked("SFXGroupHeight") Then trg.fHeight = Src.fHeight
    If IsChecked("SFXGroupScrolling") Then trg.scrolling = Src.scrolling
    
    If IsChecked("SoundsOMO") Then trg.Sound.onmouseover = Src.Sound.onmouseover
    If IsChecked("SoundsOC") Then trg.Sound.onclick = Src.Sound.onclick
    
    If IsChecked("CommandsLayout") Then trg.AlignmentStyle = Src.AlignmentStyle
    
    If IsChecked("StatusText") Then trg.WinStatus = Src.WinStatus
    
    If IsChecked("Compile") Then trg.Compile = Src.Compile
    
End Sub

Private Sub DoPasteToolbar(Src As ToolbarDef, trg As ToolbarDef)

    If IsChecked("TBToolbarStyle") Then trg.Style = Src.Style
    
    If IsChecked("TBBorderSize") Then trg.bOrder = Src.bOrder
    If IsChecked("TBBorderColor") Then trg.BorderColor = Src.BorderColor
    If IsChecked("TBBorderStyle") Then trg.BorderStyle = Src.BorderStyle
    
    If IsChecked("TBHorizontal") Then trg.ContentsMarginH = Src.ContentsMarginH
    If IsChecked("TBVertical") Then trg.ContentsMarginV = Src.ContentsMarginV
    
    If IsChecked("TBSeparation") Then trg.Separation = Src.Separation
    If IsChecked("TBJustifyHotSpots") Then trg.JustifyHotSpots = Src.JustifyHotSpots
    If IsChecked("TBBackColor") Then trg.BackColor = Src.BackColor
    If IsChecked("TBBackImage") Then trg.Image = Src.Image
    If IsChecked("TBRadius") Then trg.Radius = Src.Radius
    
    If IsChecked("TBAlignment") Then
        trg.Alignment = Src.Alignment
        trg.CustX = Src.CustX
        trg.CustY = Src.CustY
        trg.AttachTo = Src.AttachTo
        trg.AttachToAlignment = Src.AttachToAlignment
        trg.AttachToAutoResize = Src.AttachToAutoResize
    End If
    If IsChecked("TBSpanning") Then trg.Spanning = Src.Spanning
    If IsChecked("TBOHorizontal") Then trg.OffsetH = Src.OffsetH
    If IsChecked("TBOVertical") Then trg.OffsetV = Src.OffsetV
    
    If IsChecked("TBDropShadowSize") Then trg.DropShadowSize = Src.DropShadowSize
    If IsChecked("TBDropShadowColor") Then trg.DropShadowColor = Src.DropShadowColor
    If IsChecked("TBTransparency") Then trg.Transparency = Src.Transparency
    
    If IsChecked("TBFollowSB") Then
        trg.FollowHScroll = Src.FollowHScroll
        trg.FollowVScroll = Src.FollowVScroll
    End If
    
    If IsChecked("TBWidth") Then trg.Width = Src.Width
    If IsChecked("TBWidth") Then trg.Height = Src.Height
    If IsChecked("TBWidth") Then trg.Condition = Src.Condition

End Sub

Private Sub Form_Load()

    Width = GetSetting(App.EXEName, "SelCP", "WinW", 5610)
    Height = GetSetting(App.EXEName, "SelCP", "WinH", 5715)

    mnuOp.Visible = False
    
    tvProperties.ImageList = frmMain.ilIcons

    CenterForm Me
    SetupCharset Me
    LocalizeUI

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    SaveSetting App.EXEName, "SelCP", "WinW", Width
    SaveSetting App.EXEName, "SelCP", "WinH", Height

End Sub

Friend Sub Form_Resize()

    If Width < 4980 Then Width = 4980
    If Height < 4635 Then Height = 4635

    cmdOK.Move Width - 2170, Height - 900
    cmdCancel.Move cmdOK.Left + cmdOK.Width + 180, cmdOK.Top
    
    If frmPasteOptions.Visible Then
        frmPasteOptions.Top = cmdOK.Top - frmPasteOptions.Height - 160
        frmPasteOptions.Width = Width - 300
        uc3DLineSep.Move 75, frmPasteOptions.Top, Width - 300
        lblmsg.Move 690, uc3DLineSep.Top - lblmsg.Height - 30, tvProperties.Width
        tvProperties.Move 690, 60, Width - (690 + 230), lblmsg.Top - 120
        txtOp(0).Width = frmPasteOptions.Width - 370
        txtOp(1).Width = frmPasteOptions.Width - 370
        txtOp(2).Width = frmPasteOptions.Width - 370
    Else
        uc3DLineSep.Move 75, cmdOK.Top - 150, Width - 300
        lblmsg.Move 690, uc3DLineSep.Top - lblmsg.Height - 30, tvProperties.Width
        tvProperties.Move 690, 60, Width - (690 + 230), lblmsg.Top - 120
    End If
    
    imgCopy.Top = tvProperties.Top + (tvProperties.Height - imgCopy.Height) / 2
    imgPaste.Top = imgCopy.Top

End Sub

Private Sub mnuOpDesc_Click()

    If imgPaste.Visible Then
        MsgBox Mid$(tvProperties.SelectedItem.FullPath, Len(tvProperties.Nodes(1).Text) + 1) + _
                vbCrLf + vbCrLf + Replace(tvProperties.SelectedItem.tag, "%IN%", IIf(dmbClipboard.ObjSrc = docCommand, dmbClipboard.CmdContents.Name, dmbClipboard.GrpContents.Name)) + " to be pasted onto " + frmMain.tvMenus.SelectedItem.Text + ".  ", _
                vbInformation + vbOKOnly, "Property Information"
    Else
        MsgBox Mid$(tvProperties.SelectedItem.FullPath, Len(tvProperties.Nodes(1).Text) + 1) + _
                vbCrLf + vbCrLf + Replace(tvProperties.SelectedItem.tag, "%IN%", frmMain.tvMenus.SelectedItem.Text) + ".  ", _
                vbInformation + vbOKOnly, "Property Information"
    End If

End Sub

Private Sub mnuOpSelAll_Click()

    Dim nNode As Node
    
    For Each nNode In tvProperties.Nodes
        nNode.Checked = True
    Next nNode

End Sub

Private Sub mnuOpSelNone_Click()

    Dim nNode As Node
    
    For Each nNode In tvProperties.Nodes
        nNode.Checked = False
    Next nNode

End Sub

Private Sub mnuOpTreeCollapse_Click()

    Dim nNode As Node
    
    For Each nNode In tvProperties.Nodes
        If nNode.children > 0 Then
            nNode.Expanded = False
        End If
    Next nNode
    
    tvProperties.SelectedItem.EnsureVisible

End Sub

Private Sub mnuOpTreeCollapseThis_Click()

    tvProperties.SelectedItem.Expanded = False

End Sub

Private Sub mnuOpTreeExpand_Click()

    Dim nNode As Node
    
    For Each nNode In tvProperties.Nodes
        If nNode.children > 0 Then
            nNode.Expanded = True
        End If
    Next nNode
    
    tvProperties.SelectedItem.EnsureVisible

End Sub

Private Sub mnuOpTreeExpandThis_Click()
    
    tvProperties.SelectedItem.Expanded = True

End Sub

Private Sub tvProperties_Collapse(ByVal Node As MSComctlLib.Node)

    If Node.key = "ALL" Then Node.Expanded = True

End Sub

Private Sub tvProperties_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF1 Then mnuOpDesc_Click

End Sub

Private Sub tvProperties_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbRightButton Then
        If Not tvProperties.HitTest(x, y) Is Nothing Then
            tvProperties.HitTest(x, y).Selected = True
        End If
        PopupMenu mnuOp, vbRightButton, tvProperties.Left + x, tvProperties.Top + y
    End If

End Sub

Friend Sub tvProperties_NodeCheck(ByVal Node As MSComctlLib.Node)

    Dim cNode As Node
    Static Recursive As Integer
    Static UnSelEvents As Boolean
    
    If Recursive = 0 Then
        UnSelEvents = (Node.key = "ALL") Or _
                        (Node.key = "Events") Or _
                        (Node.key = "OnClick") Or _
                        (Node.key = "OnMouseOver") Or _
                        (Node.key = "OnDoubleClick")
    End If
    
    Set cNode = Node.Child
    Do Until cNode Is Nothing
        cNode.Checked = Node.Checked
        Recursive = Recursive + 1
        tvProperties_NodeCheck cNode
        Recursive = Recursive - 1
        Set cNode = cNode.Next
    Loop
    
    If Recursive = 0 Then
        Set Node = Node.parent
        Do Until Node Is Nothing
            Set cNode = Node.Child
            Node.Checked = False
            Do Until cNode Is Nothing
                If cNode.Checked Then
                    cNode.parent.Checked = True
                    Exit Do
                End If
                Set cNode = cNode.Next
            Loop
            Set Node = Node.parent
        Loop
    End If
    
    cmdOK.Enabled = False
    For Each cNode In tvProperties.Nodes
        If cNode.Checked Then
            cmdOK.Enabled = True
            Exit For
        End If
    Next cNode
    
    On Error Resume Next
    If Not IsSeparator(frmMain.tvMenus.SelectedItem.key) Then
        If Recursive = 0 Then
            If UnSelEvents Then
                UnSelEvents = False
                tvProperties.Nodes("OnClickURL").Checked = False
                tvProperties.Nodes("OnClickTargetMenu").Checked = False
                
                tvProperties.Nodes("OnMouseOverURL").Checked = False
                tvProperties.Nodes("OnMouseOverTargetMenu").Checked = False
                
                tvProperties.Nodes("OnDoubleClickURL").Checked = False
                tvProperties.Nodes("OnDoubleClickTargetMenu").Checked = False
            End If
        End If
    End If
    
End Sub

Private Sub txtOp_Click(Index As Integer)

    With opPaste(Index)
        If Not .Enabled Then Exit Sub
        .Value = True
        .SetFocus
    End With

End Sub

Friend Sub LocalizeUI()

    frmPasteOptions.caption = GetLocalizedStr(522)

    cmdAdvanced.caption = GetLocalizedStr(325) + "..."
    cmdAdvanced.Width = SetCtrlWidth(cmdAdvanced)
    
    mnuOpSel.caption = GetLocalizedStr(523)
    mnuOpSelAll.caption = GetLocalizedStr(524)
    mnuOpSelNone.caption = GetLocalizedStr(525)
    
    mnuOpTree.caption = GetLocalizedStr(526)
    mnuOpTreeExpandThis.caption = GetLocalizedStr(527)
    mnuOpTreeCollapseThis.caption = GetLocalizedStr(528)
    mnuOpTreeExpand.caption = GetLocalizedStr(529)
    mnuOpTreeCollapse.caption = GetLocalizedStr(530)
    mnuOpDesc.caption = GetLocalizedStr(531)

    cmdOK.caption = GetLocalizedStr(186)
    cmdCancel.caption = GetLocalizedStr(187)
    
    If Preferences.language <> "eng" Then
        cmdOK.Width = SetCtrlWidth(cmdOK)
        cmdCancel.Width = SetCtrlWidth(cmdCancel)
    End If

End Sub
