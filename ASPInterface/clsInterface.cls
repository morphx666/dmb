VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function LoadProject(FileName As String) As Integer

    LoadProject = Abs(LoadMenu(FileName))

End Function

Public Function Compile() As Integer

    If USER <> "DEMO" Then
        Compile = Abs(CompileProject(MenuGrps, MenuCmds, Project, Preferences, params, False))
    End If

End Function

Public Function Project_DefaultConfig(ConfigName As String)

    Project.DefaultConfig = GetConfigID(ConfigName)

End Function

Public Function Project_AddConfig(ConfigName As String, cType As ConfigTypeContants)

    With Project
        ReDim Preserve .UserConfigs(UBound(.UserConfigs) + 1)
        .UserConfigs(UBound(.UserConfigs)).Name = ConfigName
        .UserConfigs(UBound(.UserConfigs)).Type = cType
    End With

End Function

Public Function Project_Paths(Root As String, DestJS As String, DestImages As String)

    With Project.UserConfigs(Project.DefaultConfig)
        .RootWeb = Root
        .CompiledPath = DestJS
        .ImagesPath = DestImages
    End With

End Function

Public Function AddGroup(Name As String) As Integer

    On Error GoTo dspErr
    
    AddMenuGroup GetGrpParams(TemplateGroup)
    MenuGrps(UBound(MenuGrps)).Name = Name
    
    AddGroup = UBound(MenuGrps)
    
    Exit Function
    
dspErr:
    ReError

End Function

Public Function AddCommand(Name As String, Parent As Integer) As Integer

    On Error GoTo dspErr
    
    AddMenuCommand GetCmdParams(TemplateCommand)
    With MenuCmds(UBound(MenuCmds))
        .Name = Name
        .Parent = Parent
    End With

    AddCommand = UBound(MenuCmds)
    
    Exit Function
    
dspErr:
    ReError

End Function

Public Sub Group_Caption(g As Integer, Caption As String)

    MenuGrps(g).Caption = Caption

End Sub

Public Sub Group_NormalColor(g As Integer, TextColor As Long, BackgroundColor As Long)

    With MenuGrps(g)
        .nTextColor = TextColor
        .nBackColor = BackgroundColor
    End With

End Sub

Public Sub Group_OverColor(g As Integer, TextColor As Long, BackgroundColor As Long)

    With MenuGrps(g)
        .hTextColor = TextColor
        .hBackColor = BackgroundColor
    End With

End Sub

Public Sub Group_BackColor(g As Integer, Color As Long)

    With MenuGrps(g)
        .bcolor = Color
    End With

End Sub

Public Sub Group_NormalFont(g As Integer, FontStyle As String, Alignment As TextAlignmentConstants)

    ' Face, Size, Bold, Italic, Underline
    
    With MenuGrps(g).DefNormalFont
        .FontName = Split(FontStyle, ",")(0)
        .FontSize = Split(FontStyle, ",")(1)
        .FontBold = -Split(FontStyle, ",")(2)
        .FontItalic = -Split(FontStyle, ",")(3)
        .FontUnderline = -Split(FontStyle, ",")(4)
    End With
    MenuGrps(g).CaptionAlignment = Alignment

End Sub

Public Sub Group_OverFont(g As Integer, FontStyle As String)

    ' Face, Size, Bold, Italic, Underline
    
    With MenuGrps(g).DefHoverFont
        .FontName = Split(FontStyle, ",")(0)
        .FontSize = Split(FontStyle, ",")(1)
        .FontBold = -Split(FontStyle, ",")(2)
        .FontItalic = -Split(FontStyle, ",")(3)
        .FontUnderline = -Split(FontStyle, ",")(4)
    End With

End Sub

Public Sub Group_BackImage(g As Integer, BackImage As String)

    MenuGrps(g).Image = BackImage

End Sub

Public Sub Group_Frame(g As Integer, Size As Integer, LightColor As Long, DarkColor As Long)

    With MenuGrps(g)
        .FrameBorder = Size
        .lCorners = LightColor
        .dCorners = DarkColor
    End With
    
End Sub

Public Sub Group_LeftImage(g As Integer, Normal As String, Over As String, w As Long, h As Long)

    With MenuGrps(g).LeftImage
        .NormalImage = Normal
        .HoverImage = Over
        .w = w
        .h = h
    End With
    
End Sub

Public Sub Group_RightImage(g As Integer, Normal As String, Over As String, w As Long, h As Long)

    With MenuGrps(g).RightImage
        .NormalImage = Normal
        .HoverImage = Over
        .w = w
        .h = h
    End With
    
End Sub

Public Sub Group_Spacing(g As Integer, v As Integer, h As Integer, Leading As Integer)

    With MenuGrps(g)
        .ContentsMarginH = h
        .ContentsMarginV = v
        .Leading = Leading
    End With
    
End Sub

Public Sub Group_SFX(g As Integer, CmdHFXNormal As CommandsFXConstants, CmdHFXOver As CommandsFXConstants, Border As Integer, h As Integer, v As Integer)

    With MenuGrps(g)
        .CmdsFXNormal = CmdHFXNormal
        .CmdsFXOver = CmdHFXOver
        .CmdsFXSize = Border
        .CmdsMarginX = h
        .CmdsMarginY = v
    End With

End Sub

Public Sub Command_Caption(c As Integer, Caption As String)

    MenuCmds(c).Caption = Caption

End Sub

Public Sub Command_NormalColor(c As Integer, TextColor As Long, BackgroundColor As Long)

    With MenuCmds(c)
        .nTextColor = TextColor
        .nBackColor = BackgroundColor
    End With

End Sub

Public Sub Command_OverColor(c As Integer, TextColor As Long, BackgroundColor As Long)

    With MenuCmds(c)
        .hTextColor = TextColor
        .hBackColor = BackgroundColor
    End With

End Sub

Public Sub Command_NormalFont(c As Integer, FontStyle As String, Alignment As TextAlignmentConstants)

    ' Face, Size, Bold, Italic, Underline
    
    With MenuCmds(c).NormalFont
        .FontName = Split(FontStyle, ",")(0)
        .FontSize = Split(FontStyle, ",")(1)
        .FontBold = -Split(FontStyle, ",")(2)
        .FontItalic = -Split(FontStyle, ",")(3)
        .FontUnderline = -Split(FontStyle, ",")(4)
    End With
    MenuCmds(c).Alignment = Alignment

End Sub

Public Sub Command_OverFont(c As Integer, FontStyle As String)

    ' Face, Size, Bold, Italic, Underline
    
    With MenuCmds(c).HoverFont
        .FontName = Split(FontStyle, ",")(0)
        .FontSize = Split(FontStyle, ",")(1)
        .FontBold = -Split(FontStyle, ",")(2)
        .FontItalic = -Split(FontStyle, ",")(3)
        .FontUnderline = -Split(FontStyle, ",")(4)
    End With

End Sub
