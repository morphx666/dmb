Attribute VB_Name = "modGlobals"
Option Explicit

Public MenuGrps() As MenuGrp
Public MenuCmds() As MenuCmd
Public Project As ProjectDef
Public Preferences As PrgPrefs
Public params() As AddInParameter
Public FramesMode As Boolean

Public FloodPanel As New clsFlood

Public DoUNICODE As Boolean

Public AppPath As String
Public StatesPath As String
Public JSAbsPath As String
Public ImgAbsPath As String
Public HelpFile As String
Public TempPath As String
Public PreviewPath As String
Public FontCharSet As Long

Public Const LoaderCodeSTART = "<!-- DHTML Menu Builder Loader Code START -->"
Public Const LoaderCodeEND = "<!-- DHTML Menu Builder Loader Code END -->"

Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Private Type lSize
'    cx As Long
'    cy As Long
'End Type
'Private Declare Function GetTextExtentPoint32W Lib "gdi32" (ByVal hdc As Long, ByVal lpsz As Long, ByVal cbString As Long, lpSize As lSize) As Long

Public Function GetRealLocal() As ConfigDef

    GetRealLocal = IIf(Project.UserConfigs(Project.DefaultConfig).Type = ctcRemote, Project.UserConfigs(GetConfigID(Project.UserConfigs(Project.DefaultConfig).LocalInfo4RemoteConfig)), Project.UserConfigs(Project.DefaultConfig))

End Function

Public Function IsSubMenu(g As Integer) As Boolean

    Dim i As Integer
    
    For i = 1 To UBound(MenuGrps)
        If i <> g Then
            With MenuGrps(i).Actions
                IsSubMenu = (.OnClick.Type = atcCascade And .OnClick.TargetMenu = g) Or IsSubMenu
                IsSubMenu = (.OnMouseOver.Type = atcCascade And .OnMouseOver.TargetMenu = g) Or IsSubMenu
                IsSubMenu = (.OnDoubleClick.Type = atcCascade And .OnDoubleClick.TargetMenu = g) Or IsSubMenu
            End With
            If IsSubMenu Then Exit Function
        End If
    Next i
    
    If Not IsSubMenu Then
        For i = 1 To UBound(MenuCmds)
            With MenuCmds(i).Actions
                IsSubMenu = (.OnClick.Type = atcCascade And .OnClick.TargetMenu = g) Or IsSubMenu
                IsSubMenu = (.OnMouseOver.Type = atcCascade And .OnMouseOver.TargetMenu = g) Or IsSubMenu
                IsSubMenu = (.OnDoubleClick.Type = atcCascade And .OnDoubleClick.TargetMenu = g) Or IsSubMenu
            End With
            If IsSubMenu Then Exit Function
        Next i
    End If
    
End Function

Public Function Unicode2xUNI(c As String) As String

    Dim i As Integer
    Dim sHex As String
    
    If Not Project.DBCSSupport Then
        Unicode2xUNI = c
        Exit Function
    End If
    
    For i = 1 To Len(c)
        sHex = Hex(AscW(Mid(c, i, 1)))
        sHex = String$(Abs(4 - Len(sHex)), "0") + sHex
        Unicode2xUNI = Unicode2xUNI + sHex
    Next i

End Function

Public Function xUNI2Unicode(c As String) As String

    Dim i As Integer
    
    If Not Project.DBCSSupport Then
        xUNI2Unicode = c
        Exit Function
    End If
    
    For i = 1 To Len(c) Step 4
        xUNI2Unicode = xUNI2Unicode + ChrW(Val("&h" + Mid(c, i, 4)))
    Next i

End Function

Public Function xUNI2HTML(c As String) As String

    Dim i As Integer
    Dim j As Integer
    Dim a As Long
    
    If Not Project.DBCSSupport Then
        xUNI2HTML = c
        Exit Function
    End If
    
    For i = 1 To Len(c) Step 4
        a = Val("&h" + Mid(c, i, 4))
        If a <= 0 Or a > 255 Then
            xUNI2HTML = ""
            For j = 1 To Len(c) Step 4
                a = Val("&h" + Mid(c, j, 4))
                If a < 0 Then a = 65536 - a
                xUNI2HTML = xUNI2HTML + "&#" & a & ";"
            Next j
            Exit Function
        End If
        xUNI2HTML = xUNI2HTML + Chr(Val("&h" + Mid(c, i, 4)))
    Next i

End Function

Public Function FixURL(ByVal url As String) As String

    Dim RootWeb As String
    Dim p As Integer
    
    With Project.UserConfigs(Project.DefaultConfig)
        RootWeb = .RootWeb
        
        If (UsesProtocol(url) Or IsExternalLink(url)) Then
            url = SetSlashDir(url, sdFwd)
        End If
        
        Select Case .Type
            Case ctcLocal
                url = AddFileProtocol(url)
            Case ctcRemote
                If LCase(RootWeb) = LCase(Left$(url, Len(RootWeb))) And .OptmizePaths Then
                    p = InStr(InStr(url, ":") + 3, url, "/")
                    If p > 0 Then url = Mid(url, p)
                End If
                url = EscapePath(url)
            Case ctcCDROM
                If LCase(RootWeb) = LCase(Left$(url, Len(RootWeb))) And LenB(RootWeb) <> 0 Then
                    url = "%%REP%%" + GetSmartRelPath(.RootWeb, url) + GetFileName(url)
                    url = SetSlashDir(url, sdFwd)
                Else
                    If Not (UsesProtocol(url) Or IsExternalLink(url)) Then
                        If Left(url, 1) = "/" Or Left(url, 1) = "\" Then
                            url = Mid(url, 2)
                        End If
                        url = "%%REP%%" + SetSlashDir(url, sdFwd)
                    End If
                End If
        End Select
    End With
    
    url = Replace$(url, "'", Chr(30))
    url = Replace$(url, """", Chr(29))
    'url = Replace$(url, ",", Chr(28))
    
    FixURL = url

End Function

Public Function AddFileProtocol(ByVal RscPath As String) As String

    If Len(RscPath) > 2 Then
        If UCase(Left(RscPath, 1)) >= "A" And UCase(Left(RscPath, 1)) <= "Z" And Mid(RscPath, 2, 1) = ":" Then
            RscPath = CreateUrlFromPath(RscPath)
        End If
    End If

    AddFileProtocol = RscPath

End Function

Public Function EscapePath(sStr As String) As String

    If Not UsesProtocol(sStr) And Left(sStr, 7) <> "frames[" Then
        sStr = EncodeUrl(sStr)
    End If
    EscapePath = sStr

End Function

#If STANDALONE = 0 Then

Public Sub AdjustMenusAlignment(tbIndex As Integer)

    Dim i As Integer
    Dim tb As ToolbarDef
    
    tb = Project.Toolbars(tbIndex)
    
    For i = 1 To UBound(tb.Groups)
        With MenuGrps(GetIDByName(tb.Groups(i)))
            If .Actions.OnMouseOver.Type = atcCascade Then
                If tb.Style = tscHorizonal Then
                    Select Case .Alignment
                        Case gacLeftBottom, gacLeftCenter, gacLeftTop, gacRightBottom, gacRightCenter, gacRightTop
                            .Alignment = gacBottomLeft
                    End Select
                Else
                    Select Case .Alignment
                        Case gacBottomLeft, gacBottomCenter, gacBottomRight, gacTopCenter, gacTopLeft, gacTopRight
                            .Alignment = gacRightTop
                    End Select
                End If
            End If
        End With
    Next i

End Sub

Public Function RequiresImageCode() As Boolean

    Dim i As Integer
    
    If UBound(Project.Toolbars) > 0 Then
        For i = 1 To UBound(Project.Toolbars)
            If Project.Toolbars(i).Alignment = 10 And Project.Toolbars(i).AttachToAutoResize Then
                RequiresImageCode = True
                Exit For
            End If
        Next i
    Else
        RequiresImageCode = True
    End If
    
    RequiresImageCode = False
    
End Function

Public Function BelongsToToolbar(id As Integer, IsGroup As Boolean) As Integer

    Dim g As Integer
    Dim m As Integer
    Dim i As Integer
    
    If IsGroup Then
        g = id
    Else
        g = MenuCmds(id).Parent
    End If
    Do
        m = MemberOf(g)
        If m <> 0 Then Exit Do
        g = MenuCmds(SubMenuOf(g)).Parent
        If g = 0 Then Exit Do
        i = i + 1
        If i > 500 Then Exit Do
    Loop
    
    BelongsToToolbar = m

End Function

Public Function IsCascade(c As Integer) As Boolean

    With MenuCmds(c).Actions
        IsCascade = .OnClick.Type = atcCascade Or .OnDoubleClick.Type = atcCascade Or .OnMouseOver.Type = atcCascade
    End With
    
End Function

Public Function RunShellExecute(sTopic As String, sFile As Variant, sParams As Variant, sDirectory As Variant, nShowCmd As Long) As Long

   Dim hWndDesk As Long
   Dim success As Long
  
  'the desktop will be the
  'default for error messages
   hWndDesk = GetDesktopWindow()
  
  'execute the passed operation
   success = ShellExecute(hWndDesk, sTopic, sFile, sParams, sDirectory, nShowCmd)

  'This is optional. Uncomment the three lines
  'below to have the "Open With.." dialog appear
  'when the ShellExecute API call fails
  'If success = SE_ERR_NOASSOC Then
  '   Call Shell("rundll32.exe shell32.dll,OpenAs_RunDLL " & sFile, vbNormalFocus)
  'End If
  
  RunShellExecute = success
   
End Function

Public Function GetAddInDescription(AddInName As String)

    Dim ff As Integer
    Dim sStr As String
    Dim sDesc As String
    Dim fn As String
    
    On Error GoTo ExitSub
    
    ff = FreeFile
    fn = AppPath + "AddIns\" + AddInName + ".ext"
    If FileExists(fn) Then
        Open fn For Input As #ff
            Line Input #ff, sStr
            Do Until sStr = "***"
                sDesc = sDesc + sStr + vbCrLf
                Line Input #ff, sStr
            Loop
            If LenB(sDesc) <> 0 Then sDesc = Left$(sDesc, Len(sDesc) - 2)
        Close #ff
        
        GetAddInDescription = sDesc
    End If
    
    Exit Function
    
ExitSub:
    
    GetAddInDescription = "Unable to retrieve AddIn description"

End Function

Public Sub SelComboItem(sStr As String, cmb As ComboBox)

    Dim i As Integer
    
    For i = 0 To cmb.ListCount - 1
        If cmb.List(i) = sStr Then
            cmb.ListIndex = i
            Exit For
        End If
    Next i
End Sub

Public Sub CleanStatesDir()

    On Error GoTo ExitSub
    
    Dim fFile As String
 
    If Mid(StatesPath, InStrRev(StatesPath, "\", Len(StatesPath) - 1)) <> "\Default\" Then
        fFile = Dir(StatesPath + "*.*")
        Do While LenB(fFile) <> 0
            Kill StatesPath + fFile
            fFile = Dir
        Loop
    
        SetAttr StatesPath, vbNormal
        Kill StatesPath
    Else
        fFile = Dir(StatesPath + "*.dus")
        Do While LenB(fFile) <> 0
            Kill StatesPath + fFile
            fFile = Dir
        Loop
        
        fFile = Dir(StatesPath + "*.dmb")
        Do While LenB(fFile) <> 0
            Kill StatesPath + fFile
            fFile = Dir
        Loop
        
        fFile = Dir(StatesPath + "*.gif")
        Do While LenB(fFile) <> 0
            Kill StatesPath + fFile
            fFile = Dir
        Loop
        
        fFile = Dir(StatesPath + "*.jpg")
        Do While LenB(fFile) <> 0
            Kill StatesPath + fFile
            fFile = Dir
        Loop
        
        fFile = Dir(StatesPath + "*.png")
        Do While LenB(fFile) <> 0
            Kill StatesPath + fFile
            fFile = Dir
        Loop
    End If
    
ExitSub:

End Sub

Public Sub CleanPreviewDir(DelAll As Boolean)

    On Error Resume Next
    
    Dim ext As String
    Dim fFile As String
    Dim pPath As String
    
    If LenB(PreviewPath) = 0 Then Exit Sub
    
    pPath = PreviewPath
    
    fFile = Dir(pPath + "*.*")
    Do While LenB(fFile) <> 0
        If InStr(fFile, ".") Then ext = Split(fFile, ".")(1)
        If DelAll Then
            SetAttr pPath + fFile, vbNormal
            Kill pPath + fFile
        ElseIf ext = "html" Or ext = "js" Or ext = "txt" Then
            SetAttr pPath + fFile, vbNormal
            Kill pPath + fFile
        End If
        
        fFile = Dir
    Loop

End Sub

Private Function GetAllColors() As Long()

    Dim c() As Long
    Dim i As Integer
    Dim k As Integer
    Dim n As Integer
    
    On Error Resume Next
    
    ReDim c(1 To UBound(MenuGrps) * 12 + UBound(MenuCmds) * 4 + UBound(Project.Toolbars) * 2)
    
    n = 4
    For i = 1 To UBound(MenuCmds)
        With MenuCmds(i)
            c(1 + (i - 1) * n) = .nTextColor
            c(2 + (i - 1) * n) = .nBackColor
            c(3 + (i - 1) * n) = .hTextColor
            c(4 + (i - 1) * n) = .hBackColor
        End With
    Next i
    
    k = 1 + (i - 1) * n
    For i = 1 To UBound(Project.Toolbars)
        c(k) = Project.Toolbars(i).BackColor
        c(k + 1) = Project.Toolbars(i).BorderColor
        k = k + 2
    Next i
    
    n = 10
    For i = 1 To UBound(MenuGrps)
        With MenuGrps(i)
            c(k + 1 + (i - 1) * n) = .bColor
            c(k + 2 + (i - 1) * n) = .Corners.leftCorner
            c(k + 3 + (i - 1) * n) = .Corners.rightCorner
            c(k + 4 + (i - 1) * n) = .nTextColor
            c(k + 5 + (i - 1) * n) = .nBackColor
            c(k + 6 + (i - 1) * n) = .hTextColor
            c(k + 7 + (i - 1) * n) = .hBackColor
            c(k + 8 + (i - 1) * n) = .Corners.topCorner
            c(k + 9 + (i - 1) * n) = .Corners.bottomCorner
            c(k + 10 + (i - 1) * n) = .DropShadowColor
            c(k + 11 + (i - 1) * n) = .CmdsFXnColor
            c(k + 12 + (i - 1) * n) = .CmdsFXhColor
        End With
    Next i
    
    GetAllColors = c

End Function

Public Sub BuildUsedColorsArray()

    Dim c() As Long
    Dim k As Integer
    Dim j As Integer
    Dim Ok2Add As Boolean
    
    c = GetAllColors
    ReDim UsedColors(0)
    For j = 1 To UBound(c)
        Ok2Add = True
        For k = 1 To UBound(UsedColors)
            If UsedColors(k) = c(j) Then
                Ok2Add = False
                Exit For
            End If
        Next k
        If Ok2Add Then
            ReDim Preserve UsedColors(UBound(UsedColors) + 1)
            UsedColors(UBound(UsedColors)) = c(j)
        End If
    Next j
    
End Sub

Public Sub SetupCharset(rForm As Form)

    Dim ctrl As Control
    Dim FontObj As Object
    
    On Error Resume Next
    
    For Each ctrl In rForm.Controls
        Err.Clear
        Set FontObj = ctrl.Font
        If Err.Number = 0 Then
            FontObj.Charset = FontCharSet
        End If
    Next ctrl

End Sub

#End If

Public Function SetSlashDir(ByVal str As String, d As SlashDir) As String

    Select Case d
        Case sdFwd
            str = Replace(str, "\", "/")
        Case sdBack
            str = Replace(str, "/", "\")
    End Select
    
    SetSlashDir = str

End Function

Public Sub UpdateItemsLinks()

    Dim i As Integer
    
    On Error GoTo ExitSub
    
    For i = 1 To UBound(MenuCmds)
        MenuCmds(i).Actions.OnMouseOver.url = ConvertPath(MenuCmds(i).Actions.OnMouseOver.url)
        MenuCmds(i).Actions.OnClick.url = ConvertPath(MenuCmds(i).Actions.OnClick.url)
        MenuCmds(i).Actions.OnDoubleClick.url = ConvertPath(MenuCmds(i).Actions.OnDoubleClick.url)
    Next i
    
    For i = 1 To UBound(MenuGrps)
        MenuGrps(i).Actions.OnMouseOver.url = ConvertPath(MenuGrps(i).Actions.OnMouseOver.url)
        MenuGrps(i).Actions.OnClick.url = ConvertPath(MenuGrps(i).Actions.OnClick.url)
        MenuGrps(i).Actions.OnDoubleClick.url = ConvertPath(MenuGrps(i).Actions.OnDoubleClick.url)
    Next i
    
ExitSub:

End Sub

Public Function GetSmartRelPath(FromDir As String, ToDir As String) As String

    Dim RelPath As String
    Dim FromDirs() As String
    Dim ToDirs() As String
    Dim i As Integer
    Dim p As Integer
    
    Dim fdl As Integer
    Dim tdl As Integer

    On Error GoTo chkError
    
    If LenB(FromDir) = 0 Or LenB(ToDir) = 0 Then Exit Function
    If LCase(FromDir) = LCase(ToDir) Then Exit Function
    
    FromDir = RemoveDoubleSlashes(FromDir)
    ToDir = RemoveDoubleSlashes(ToDir)

    FromDirs = Split(GetFilePath(FromDir), "\"): fdl = UBound(FromDirs) - 1
    ToDirs = Split(GetFilePath(ToDir), "\"): tdl = UBound(ToDirs) - 1
    ReDim Preserve FromDirs(fdl)
    ReDim Preserve ToDirs(tdl)

    If tdl < fdl Then
        For i = 0 To fdl
            If i > tdl Or p > 0 Then
                RelPath = RelPath + "..\"
                If p = 0 Then p = i
            ElseIf LCase(ToDirs(i)) <> LCase(FromDirs(i)) And p = 0 Then
                RelPath = RelPath + "..\"
                p = i
            End If
        Next i
    ElseIf tdl >= fdl Then
        For i = 0 To tdl
            If i > fdl Then
                If p = 0 Then p = i
            ElseIf LCase(ToDirs(i)) <> LCase(FromDirs(i)) Then
                RelPath = RelPath + "..\"
                If p = 0 Then p = i
            End If
        Next i
    End If

    If p > 0 Then
        For i = p To tdl
            RelPath = RelPath + ToDirs(i) + "\"
        Next i
    End If

    GetSmartRelPath = RelPath

    Exit Function

chkError:
    MsgBox "Your Project Properties contains invalid information about the Root Web or the Folder to Store Compiled Files." + vbCrLf + "Please correct this problem before trying to add links to your commands.", vbInformation + vbOKOnly, "Invalid Project Properties"

End Function

Public Function GetIDByName(ByVal gName As String) As Integer

    Dim i As Integer

    For i = 1 To UBound(MenuGrps)
        If MenuGrps(i).Name = gName Then
            GetIDByName = i
            Exit Function
        End If
    Next i
    
    For i = 1 To UBound(MenuCmds)
        If MenuCmds(i).Name = gName Then
            GetIDByName = i
            Exit Function
        End If
    Next i

End Function

Public Function GetSingleCommandHeight(c As Integer) As Integer

    Dim accHeight As Integer
    Dim ImgHeight As Integer
    Dim numCommands As Integer
    Dim i As Integer
    Dim HasCaption As Boolean
    Dim IsLast As Boolean
    Dim avHeight As Integer
    
    With MenuCmds(c)
        If .Name <> "[SEP]" Then
            If MenuGrps(.Parent).fHeight <> 0 And MenuGrps(.Parent).AlignmentStyle = ascVertical Then
                For i = 1 To UBound(MenuCmds)
                    If MenuCmds(i).Parent = .Parent Then numCommands = numCommands + 1
                Next i
                IsLast = ((i - 1) = c)
                
                avHeight = GetDivHeight(.Parent)
                With MenuGrps(.Parent)
                    avHeight = avHeight - .Leading * (numCommands - 1) - 2 * .FrameBorder
                    accHeight = avHeight \ numCommands
                    accHeight = accHeight - 2 * MenuCmds(c).CmdsFXSize
                    
                    If IsLast Then
                        accHeight = avHeight - (accHeight + .Leading + 2 * MenuCmds(c).CmdsFXSize) * (numCommands - 1)
                    End If
                End With
            Else
                If LenB(.Caption) <> 0 Or LenB(.BackImage.NormalImage) = 0 Then
                    accHeight = GetCmdTextHeight(c)
                    
                    ImgHeight = IIf(.LeftImage.h > .RightImage.h, .LeftImage.h, .RightImage.h)
                    If ImgHeight > accHeight Then accHeight = ImgHeight
                    
                    accHeight = accHeight + 2 * .CmdsMarginY
                Else
                    For i = 1 To UBound(MenuCmds)
                        If MenuCmds(i).Parent = .Parent And MenuCmds(i).Compile Then
                            HasCaption = (LenB(MenuCmds(i).Caption) <> 0)
                            If HasCaption Then Exit For
                        End If
                    Next i
                    If Not HasCaption Or MenuGrps(.Parent).AlignmentStyle = ascVertical Then
                        Set frmMain.picRsc.Picture = LoadPictureRes(.BackImage.NormalImage)
                        accHeight = frmMain.picRsc.Height / Screen.TwipsPerPixelY - 2 * .CmdsFXSize
                    End If
                End If
            End If
            
            applyCroppingHeight .BackImage, accHeight
        Else
            accHeight = Preferences.SepHeight
        End If
    End With
    
    GetSingleCommandHeight = accHeight
    
End Function

Public Function GetSingleCommandWidth(c As Integer) As Integer

    Dim tmpWidth As Integer
    Dim accWidth As Integer
    Dim numCommands As Integer
    Dim i As Integer
    Dim HasCaption As Boolean
    Dim IsLast As Boolean
    Dim avWidth As Integer
    
    With MenuCmds(c)
        If .Name <> "[SEP]" Then
            If MenuGrps(.Parent).fWidth <> 0 And MenuGrps(.Parent).AlignmentStyle = ascHorizontal Then
                For i = 1 To UBound(MenuCmds)
                    If MenuCmds(i).Parent = .Parent Then numCommands = numCommands + 1
                Next i
                IsLast = ((i - 1) = c)
                
                avWidth = GetDivWidth(.Parent)
                With MenuGrps(.Parent)
                    avWidth = avWidth - .Leading * (numCommands - 1) - 2 * .FrameBorder
                    tmpWidth = avWidth \ numCommands
                    tmpWidth = tmpWidth - 2 * MenuCmds(c).CmdsFXSize
                    
                    If IsLast Then
                        tmpWidth = avWidth - (tmpWidth + .Leading + 2 * MenuCmds(c).CmdsFXSize) * (numCommands - 1)
                    End If
                End With
            Else
                If LenB(.Caption) <> 0 Or LenB(.BackImage.NormalImage) = 0 Then
                
                    tmpWidth = GetCmdTextWidth(c)
                    
                    If LenB(.LeftImage.NormalImage) <> 0 Then
                        tmpWidth = tmpWidth + .LeftImage.w + Preferences.ImgSpace
                    End If
                    If LenB(.RightImage.NormalImage) <> 0 Then
                        tmpWidth = tmpWidth + .RightImage.w + 2 * Preferences.ImgSpace
                    End If
                    tmpWidth = tmpWidth + 2 * .CmdsMarginX + 2 * .CmdsFXSize
                    
                    If tmpWidth > accWidth Then accWidth = tmpWidth
                Else
                    For i = 1 To UBound(MenuCmds)
                        If MenuCmds(i).Parent = .Parent And MenuCmds(i).Compile Then
                            HasCaption = (LenB(MenuCmds(i).Caption) <> 0)
                            If HasCaption Then Exit For
                        End If
                    Next i
                    If Not HasCaption Or MenuGrps(.Parent).AlignmentStyle = ascHorizontal Then
                        If LenB(.BackImage.NormalImage) <> 0 Then
                            Set frmMain.picRsc.Picture = LoadPictureRes(.BackImage.NormalImage)
                            'tmpWidth = frmMain.picRsc.Width / Screen.TwipsPerPixelX - 2 * MenuGrps(.Parent).CmdsFXSize
                            tmpWidth = frmMain.picRsc.Width / Screen.TwipsPerPixelX - 2 * .CmdsFXSize
                        End If
                    End If
                End If
            End If
            
            applyCroppingWidth .BackImage, tmpWidth
        Else
            tmpWidth = Preferences.SepHeight + 2
        End If
    End With
    
    GetSingleCommandWidth = tmpWidth
    
End Function

Public Function GetCmdTextHeight(c As Integer) As Integer
    
    Dim sn As Integer
    Dim sO As Integer
    Dim cp As String
    
    With MenuCmds(c)
        cp = ParseCaptionHTML(.Caption)
        sn = GetTextSize(cp, .NormalFont)(2)
        sO = GetTextSize(cp, .HoverFont)(2)
    End With
    GetCmdTextHeight = IIf(sn >= sO, sn, sO)
    
End Function

Public Function GetCmdTextWidth(c As Integer) As Integer
    
    Dim sn As Integer
    Dim sO As Integer
    Dim cp As String
    
    With MenuCmds(c)
        cp = ParseCaptionHTML(.Caption)
        sn = GetTextSize(cp, .NormalFont, , , False)(1)
        sO = GetTextSize(cp, .HoverFont, , , False)(1)
    End With
    GetCmdTextWidth = IIf(sn >= sO, sn, sO)
    
End Function

Public Function GetGrpTextHeight(g As Integer) As Integer

    Dim sn As Integer
    Dim sO As Integer
    Dim cp As String
    
    With MenuGrps(g)
        cp = ParseCaptionHTML(.Caption)
        sn = GetTextSize(cp, .DefNormalFont, , , False)(2)
        sO = GetTextSize(cp, .DefHoverFont, , , False)(2)
    End With
    GetGrpTextHeight = IIf(sn >= sO, sn, sO)
    
End Function

Public Function GetGrpTextWidth(g As Integer) As Integer

    Dim sn As Integer
    Dim sO As Integer
    Dim cp As String
    
    With MenuGrps(g)
        cp = ParseCaptionHTML(.Caption)
        sn = GetTextSize(cp, .DefNormalFont)(1)
        sO = GetTextSize(cp, .DefHoverFont)(1)
    End With
    GetGrpTextWidth = IIf(sn >= sO, sn, sO)
    
End Function

Public Function GetTextSize(ByVal Text As String, Optional sFont As Variant, Optional vbObj As Object, Optional ScaleMode As Integer = vbTwips, Optional DoRemoveHTMLCode As Boolean = True) As Integer()

    'Dim picObj As PictureBox
    Dim Size() As Integer
    Dim idx As Integer
    
    ReDim Size(1 To 2)
    
    If frmMain Is Nothing Then Exit Function
    
    On Error Resume Next
    If LenB(Text) <> 0 Then
'        Do
'            idx = idx + 1
'            Err.Clear
'            Set picObj = frmMain.Controls.Add("VB.PictureBox", "picObj" & idx)
'            If Err.number = 0 Then Exit Do
'        Loop
        
        'With picObj
        With frmMain.picObj1
            .Font.Charset = FontCharSet
            If vbObj Is Nothing Then
                .ScaleMode = vbPixels
                With .Font
                    .Name = sFont.FontName
                    .Size = px2pt(sFont.FontSize)
                    .Bold = sFont.FontBold
                    .Italic = sFont.FontItalic
                    .Underline = sFont.FontUnderline
                End With
            Else
                .ScaleMode = ScaleMode
                With .Font
                    .Name = vbObj.FontName
                    .Size = vbObj.FontSize
                    .Bold = vbObj.FontBold
                    .Italic = vbObj.FontItalic
                    .Underline = vbObj.FontUnderline
                End With
            End If
            
            'Dim udtSize As lSize
            'GetTextExtentPoint32W .hdc, StrPtr(Text), Len(Text), udtSize
            'Size(1) = udtSize.cx
            'Size(2) = udtSize.cy
            
            If DoRemoveHTMLCode Then Text = RemoveHTMLCode(Text)
            If DoUNICODE Then
                Text = Replace(StrConv(xUNI2Unicode(Text), vbUnicode), vbNullChar, "")
            Else
                Text = xUNI2Unicode(Text)
            End If
            
            Text = RemoveUnicodeChars(Text)
            
            Size(1) = .TextWidth(Text) - 10 * ((15 / Screen.TwipsPerPixelX) - 1)
            Size(2) = .TextHeight(Text) - 10 * ((15 / Screen.TwipsPerPixelY) - 1)
            
            If Size(1) < 0 Then Size(1) = 0
            If Size(2) < 0 Then Size(2) = 0
        End With
        
'        Do
'            Err.Clear
'            frmMain.Controls.Remove "picObj" & idx
'            Set picObj = Nothing
'            Set picObj = frmMain.Controls("picObj" & idx)
'            If Not picObj Is Nothing Then Stop
'        Loop Until Err.number <> 0
    End If
    
    GetTextSize = Size

End Function

Private Function RemoveUnicodeChars(ByVal Text As String) As String
'    Dim s As String
'    Dim m As String
'    Dim i As Integer
'
'    For i = 1 To Len(Text)
'        m = Mid(Text, i, 1)
'        If Asc(m) >= 9 Then s = s + m
'    Next i
'
'    RemoveUnicodeChars = s

    Dim i As Integer
    
    For i = 0 To 8
        Text = Replace(Text, Chr(i), "")
    Next i
    
    RemoveUnicodeChars = Text
End Function

Private Sub applyCroppingWidth(ByRef img As tImage, ByRef w As Integer)

    With img
        If LenB(.NormalImage) <> 0 Then
            If .AllowCrop = False And .w > w Then w = .w
        End If
    End With

End Sub

Private Sub applyCroppingHeight(ByRef img As tImage, ByRef h As Integer)

    With img
        If LenB(.NormalImage) <> 0 Then
            If .AllowCrop = False And .h > h Then h = .h
        End If
    End With

End Sub

Public Function RemoveHTMLCode(ByVal sCode As String) As String

    Dim p1 As Integer
    Dim p2 As Integer
    Dim p3 As Integer
    Dim p4 As Integer
    Dim s As String
    Dim ss As String
    Dim k As Integer
    
    On Error GoTo AbortFcn
    
    s = sCode
    
    While (InStr(s, "<") > 0) And (InStr(s, ">") > 0) And k < 100
        k = k + 1
        ss = ""
        p1 = 1
        Do
            p1 = InStr(p1, s, "<")
            If p1 > 0 Then
                If p1 > 1 Then ss = ss + Left(s, p1 - 1)
                p2 = InStr(p1, s, ">")
                p3 = InStr(p1, s, " ")
                If p2 = 0 Then Exit Do
                If p3 < p2 And p3 <> 0 Then
                    p4 = p3
                    s = Replace(s, "</" + Mid(s, p1 + 1, p4 - p1 - 1) + ">", "")
                Else
                    p4 = p2
                    s = Replace(s, "</" + Mid(s, p1 + 1, p4 - p1), "")
                End If
                s = Mid(s, p2 + 1)
            Else
                Exit Do
            End If
        Loop
        s = ss + s
    Wend
    
    RemoveHTMLCode = s
    
    Exit Function
    
AbortFcn:

    RemoveHTMLCode = sCode

End Function

Public Function ParseCaptionHTML(ByVal sStr As String) As String

    Dim p1 As Long
    Dim p2 As Long
    Dim v As Long
    
    On Error Resume Next
    
    p1 = InStr(sStr, "&#")
    If p1 > 0 Then
    
        Do
            Dim p3 As Long
            Dim p As String
            p2 = 0
            For p3 = p1 + 2 To Len(sStr)
                p = Mid(sStr, p3, 1)
                If p = ";" Then
                    p2 = p3
                    Exit For
                Else
                    If p < "0" Or p > "9" Then Exit For
                End If
            Next p3
            If p2 > 0 Then
                v = Abs(Val(Mid(sStr, p1 + 2, p2 - p1 - 2)))
                If v > 255 Then
                    If v > &H207F Then
                        sStr = Replace(sStr, "&#" & v & ";", "— ")
                    Else
                        sStr = Left(sStr, p1 - 1) + ChrW(v) + Mid(sStr, p2 + 1)
                    End If
                Else
                    sStr = Left(sStr, p1 - 1) + Chr(v) + Mid(sStr, p2 + 1)
                End If
            End If
            p1 = InStr(p1 + 1, sStr, "&#")
            If p1 = 0 Then Exit Do
        Loop
        
'        p2 = InStr(p1, sStr, ";")
'        Do While p2 > p1 And p1 <> 0
'            v = Abs(Val(Mid(sStr, p1 + 2, p2 - p1 - 2)))
'            If v > 255 Then
'                If v > &H207F Then
'                    sStr = Replace(sStr, "&#" & v & ";", "— ")
'                Else
'                    sStr = Left(sStr, p1 - 1) + ChrW(v) + Mid(sStr, p2 + 1)
'                End If
'            Else
'                sStr = Left(sStr, p1 - 1) + Chr(v) + Mid(sStr, p2 + 1)
'            End If
'            p1 = InStr(sStr, "&#")
'            If p1 > 0 Then p2 = InStr(p1, sStr, ";")
'        Loop
    End If
    
    sStr = Replace(sStr, "&amp;", "&")
    sStr = Replace(sStr, "&nbsp;", " ")
    sStr = Replace(sStr, "&cent;", "¢")
    sStr = Replace(sStr, "&pound;", "£")
    sStr = Replace(sStr, "&yen;", "¥")
    sStr = Replace(sStr, "&frac14;", "¼")
    sStr = Replace(sStr, "&frac12;", "½")
    sStr = Replace(sStr, "&frac34;", "¾")
    sStr = Replace(sStr, "&plusmn;", "±")
    sStr = Replace(sStr, "&reg;", "®")
    sStr = Replace(sStr, "&deg;", "°")
    sStr = Replace(sStr, "&copy;", "©")
    sStr = Replace(sStr, "&laquo;", "«")
    sStr = Replace(sStr, "&raquo;", "»")
    sStr = Replace(sStr, "&micro;", "µ")
    sStr = Replace(sStr, "&iquest;", "¿")
    sStr = Replace(sStr, "&middot;", "·")
    sStr = Replace(sStr, "&quot;", """")
    sStr = Replace(sStr, "&Oslash;", "Ø")
    sStr = Replace(sStr, "<br>", vbCrLf)
    sStr = Replace(sStr, "<BR>", vbCrLf)
    
    sStr = Replace(sStr, "&agrave;", "à")
    sStr = Replace(sStr, "&aacute;", "á")
    sStr = Replace(sStr, "&acirc;", "â")
    sStr = Replace(sStr, "&atilde;", "ã")
    sStr = Replace(sStr, "&auml;", "ä")
    sStr = Replace(sStr, "&aring;", "å")
    sStr = Replace(sStr, "&aelig;", "æ")
    
    sStr = Replace(sStr, "&egrave;", "è")
    sStr = Replace(sStr, "&eacute;", "é")
    sStr = Replace(sStr, "&ecirc;", "ê")
    sStr = Replace(sStr, "&euml;", "ë")
    
    sStr = Replace(sStr, "&igrave;", "ì")
    sStr = Replace(sStr, "&iacute;", "í")
    sStr = Replace(sStr, "&icirc;", "î")
    sStr = Replace(sStr, "&iuml;", "ï")
    
    sStr = Replace(sStr, "&ograve;", "ò")
    sStr = Replace(sStr, "&oacute;", "ó")
    sStr = Replace(sStr, "&ocirc;", "ô")
    sStr = Replace(sStr, "&otilde;", "õ")
    sStr = Replace(sStr, "&ouml;", "ö")
    sStr = Replace(sStr, "&oslash;", "ø")
    
    sStr = Replace(sStr, "&ugrave;", "ù")
    sStr = Replace(sStr, "&uacute;", "ú")
    sStr = Replace(sStr, "&ucirc;", "û")
    sStr = Replace(sStr, "&uuml;", "ü")
    
    sStr = Replace(sStr, "&iexcl;", "¡")
    sStr = Replace(sStr, "&brvbar;", "¦")
    sStr = Replace(sStr, "&sect;", "§")
    sStr = Replace(sStr, "&uml;", "¨")
    sStr = Replace(sStr, "&ordf;", "ª")
    sStr = Replace(sStr, "&not;", "¬")
    sStr = Replace(sStr, "&shy;", "­")
    sStr = Replace(sStr, "&macr;", "¯")
    sStr = Replace(sStr, "&sup2;", "²")
    sStr = Replace(sStr, "&sup3;", "³")
    sStr = Replace(sStr, "&acute;", "´")
    sStr = Replace(sStr, "&para;", "¶")
    sStr = Replace(sStr, "&cedil;", "¸")
    sStr = Replace(sStr, "&sup1;", "¹")
    sStr = Replace(sStr, "&ordm;", "º")
    
    sStr = Replace(sStr, "&Agrave;", "À")
    sStr = Replace(sStr, "&Aacute;", "Á")
    sStr = Replace(sStr, "&Acirc;", "Â")
    sStr = Replace(sStr, "&Atilde;", "Ã")
    sStr = Replace(sStr, "&Auml;", "Ä")
    sStr = Replace(sStr, "&Aring;", "Å")
    sStr = Replace(sStr, "&AElig;", "Æ")
    
    sStr = Replace(sStr, "&Ccedil;", "Ç")
    
    sStr = Replace(sStr, "&Egrave;", "È")
    sStr = Replace(sStr, "&Eacute;", "É")
    sStr = Replace(sStr, "&Ecirc;", "Ê")
    sStr = Replace(sStr, "&Euml;", "Ë")
    
    sStr = Replace(sStr, "&Igrave;", "Ì")
    sStr = Replace(sStr, "&Iacute;", "Í")
    sStr = Replace(sStr, "&Icirc;", "Î")
    sStr = Replace(sStr, "&Iuml;", "Ï")
    
    sStr = Replace(sStr, "&ETH;", "Ð")
    sStr = Replace(sStr, "&eth;", "ð")
    
    sStr = Replace(sStr, "&Ntilde;", "Ñ")
    sStr = Replace(sStr, "&ntilde;", "ñ")
    
    sStr = Replace(sStr, "&Ograve;", "Ò")
    sStr = Replace(sStr, "&Oacute;", "Ó")
    sStr = Replace(sStr, "&Ocirc;", "Ô")
    sStr = Replace(sStr, "&Otilde;", "Õ")
    sStr = Replace(sStr, "&Ouml;", "Ö")
    sStr = Replace(sStr, "&Oslash;", "Ø")
    
    sStr = Replace(sStr, "&Ugrave;", "Ù")
    sStr = Replace(sStr, "&Uacute;", "Ú")
    sStr = Replace(sStr, "&Ucirc;", "Û")
    sStr = Replace(sStr, "&Uuml;", "Ü")
    
    sStr = Replace(sStr, "&Yacute;", "Ý")
    sStr = Replace(sStr, "&yacute;", "ý")
    sStr = Replace(sStr, "&yuml;", "ÿ")
    
    sStr = Replace(sStr, "&THORN;", "Þ")
    sStr = Replace(sStr, "&thorn;", "þ")
    
    sStr = Replace(sStr, "&szlig;", "ß")
    sStr = Replace(sStr, "&divide;", "÷")
    
    sStr = RemoveHTMLCode(sStr)
    
    sStr = Replace(sStr, "&lt;", "<")
    sStr = Replace(sStr, "&gt;", ">")
    
    ParseCaptionHTML = sStr

End Function

Private Function IsInTB(g As Integer, t As Integer) As Boolean

    IsInTB = (InStr(Join(Project.Toolbars(t).Groups, "|") + "|", "|" + MenuGrps(g).Name + "|") > 0)

End Function

Public Function GetTBWidth(t As Integer, Optional IgnoreSettings As Boolean) As Integer()

    Dim g As Integer
    Dim numHotSpots As Integer
    Dim Wider As Integer
    Dim totalWidth As Integer
    Dim r(1 To 2) As Integer
    Dim oWidth As Integer
    Dim cTB As ToolbarDef
    Dim hpw As Integer
    
    cTB = Project.Toolbars(t)
    
    With cTB
        oWidth = .Width
        If IgnoreSettings Then .Width = 0
    
        If .Width < 0 Then
            For g = 1 To UBound(MenuGrps)
                If MenuGrps(g).Compile Then
                    If IsInTB(g, t) Then numHotSpots = numHotSpots + 1
                End If
            Next g
            totalWidth = Abs(.Width)
            If .Style = tscHorizonal Then
                If numHotSpots > 0 Then Wider = Int((totalWidth - 2 * (.Border + .ContentsMarginH) - (numHotSpots - 1) * .Separation) / numHotSpots)
            Else
                Wider = totalWidth - 2 * (.Border + .ContentsMarginH) '- (numHotSpots - 1) * .Separation
            End If
        Else
            For g = 1 To UBound(MenuGrps)
                If MenuGrps(g).Compile Then
                    If IsInTB(g, t) Then
                        numHotSpots = numHotSpots + 1
                        
                        hpw = GetHotSpotWidth(t, g, IgnoreSettings)
                        totalWidth = totalWidth + hpw
                        If hpw > Wider Then Wider = hpw
                    End If
                End If
            Next g
            If .Style = tscVertical Then
                totalWidth = Wider
            Else
                If .JustifyHotSpots Then
                    totalWidth = Wider * numHotSpots + .Separation * (numHotSpots - 1)
                Else
                    totalWidth = totalWidth + .Separation * (numHotSpots - 1)
                End If
            End If
            totalWidth = totalWidth + 2 * (.Border + .ContentsMarginH)
        End If
        
        .Width = oWidth
    End With
    
    r(1) = totalWidth
    r(2) = Wider
    GetTBWidth = r

End Function

Public Function GetTBHeight(t As Integer, Optional IgnoreSettings As Boolean) As Integer()

    Dim g As Integer
    Dim numHotSpots As Integer
    Dim Taller As Integer
    Dim totalHeight As Integer
    Dim r(1 To 2) As Integer
    Dim oHeight As Integer
    Dim cTB As ToolbarDef
    Dim hph As Integer
    
    cTB = Project.Toolbars(t)
    
    With cTB
        oHeight = .Height
        If IgnoreSettings Then .Height = 0
        
        If .Height < 0 Then
            For g = 1 To UBound(MenuGrps)
                If MenuGrps(g).Compile Then
                    If IsInTB(g, t) Then numHotSpots = numHotSpots + 1
                End If
            Next g
            totalHeight = Abs(.Height)
            If .Style = tscVertical Then
                If numHotSpots > 0 Then Taller = Int((totalHeight - (2 * .Border + .ContentsMarginV) - (numHotSpots - 1) * .Separation) / numHotSpots)
            Else
                Taller = totalHeight - (2 * .Border + .ContentsMarginV) '- (numHotSpots - 1) * .Separation
            End If
        Else
            For g = 1 To UBound(MenuGrps)
                If MenuGrps(g).Compile Then
                    If IsInTB(g, t) Then
                        numHotSpots = numHotSpots + 1
                        
                        hph = GetHotSpotHeight(t, g, IgnoreSettings)
                        totalHeight = totalHeight + hph
                        If hph > Taller Then Taller = hph
                    End If
                End If
            Next g
            If .Style = tscHorizonal Then
                totalHeight = Taller
            Else
                If .JustifyHotSpots Then
                    totalHeight = (Taller * numHotSpots) + .Separation * (numHotSpots - 1)
                Else
                    totalHeight = totalHeight + .Separation * (numHotSpots - 1)
                End If
            End If
            totalHeight = totalHeight + 2 * (.Border + .ContentsMarginV)
        End If
        
        .Height = oHeight
    End With
    
    r(1) = totalHeight
    r(2) = Taller
    GetTBHeight = r

End Function

Public Function GetDivWidth(g As Integer) As Integer

    Dim accWidth As Integer
    Dim tmpWidth As Integer
    Dim c As Integer
    
    Select Case MenuGrps(g).fWidth
        Case -1
            Set frmMain.picRsc.Picture = LoadPictureRes(MenuGrps(g).Image)
            GetDivWidth = frmMain.picRsc.Width / Screen.TwipsPerPixelX
            Exit Function
        Case Is > 0
            GetDivWidth = MenuGrps(g).fWidth
            Exit Function
    End Select
    
    Select Case MenuGrps(g).AlignmentStyle
        Case ascVertical
            For c = 1 To UBound(MenuCmds)
                With MenuCmds(c)
                    If .Parent = g And .Name <> "[SEP]" And .Compile Then
                        tmpWidth = GetSingleCommandWidth(c)
                        If tmpWidth > accWidth Then accWidth = tmpWidth
                    End If
                End With
            Next c
        Case ascHorizontal
            For c = 1 To UBound(MenuCmds)
                With MenuCmds(c)
                    If .Parent = g And .Compile Then
                        accWidth = accWidth + GetSingleCommandWidth(c) + MenuGrps(.Parent).Leading
                        'If .Name = "[SEP]" Then accWidth = accWidth - 2
                    End If
                End With
            Next c
            accWidth = accWidth - MenuGrps(g).Leading
    End Select
    GetDivWidth = accWidth + CalcFrameBorder(g, ascHorizontal) + 2 * MenuGrps(g).ContentsMarginH
    
End Function

Public Function CalcFrameBorder(g As Integer, wh As AlignmentStyleConstants) As Integer

    Dim acc As Integer

    With MenuGrps(g)
        If .FrameBorder > 0 Then
            Select Case wh
                Case ascHorizontal
                    acc = Abs(.Corners.leftCorner <> -2) + Abs(.Corners.rightCorner <> -2)
                Case ascVertical
                    acc = Abs(.Corners.topCorner <> -2) + Abs(.Corners.bottomCorner <> -2)
            End Select
            acc = acc * .FrameBorder
        End If
    End With
    
    CalcFrameBorder = acc

End Function

Public Function GetDivHeight(g As Integer, Optional LimitCmd As Integer) As Integer

    Dim accHeight As Integer
    Dim tmpHeight As Integer
    Dim c As Integer
    Dim nc As Integer
    
    Select Case MenuGrps(g).fHeight
        Case -1
            Set frmMain.picRsc.Picture = LoadPictureRes(MenuGrps(g).Image)
            GetDivHeight = frmMain.picRsc.Height / Screen.TwipsPerPixelY
            Exit Function
        Case Is > 0
            GetDivHeight = MenuGrps(g).fHeight
            Exit Function
    End Select
    
    Select Case MenuGrps(g).AlignmentStyle
        Case ascVertical
            accHeight = CalcFrameBorder(g, ascVertical) + 2 * MenuGrps(g).ContentsMarginV
            For c = 1 To UBound(MenuCmds)
                With MenuCmds(c)
                    If .Parent = g And .Compile Then
                        nc = nc + 1
                        accHeight = accHeight + GetSingleCommandHeight(c)
                        If .Name <> "[SEP]" Then
                            accHeight = accHeight + MenuGrps(g).Leading + 2 * .CmdsFXSize
                        Else
                            accHeight = accHeight + MenuGrps(g).Leading
                        End If
                    End If
                    If nc = LimitCmd And LimitCmd > 0 Then Exit For
                End With
            Next c
            
            GetDivHeight = accHeight - MenuGrps(g).Leading
        Case ascHorizontal
            For c = 1 To UBound(MenuCmds)
                With MenuCmds(c)
                    If .Parent = g And .Name <> "[SEP]" And .Compile Then
                        tmpHeight = GetSingleCommandHeight(c)
                        If MenuCmds(c).Name <> "[SEP]" Then
                            tmpHeight = tmpHeight + 2 * MenuCmds(c).CmdsFXSize
                        End If
                        If tmpHeight > accHeight Then accHeight = tmpHeight
                    End If
                End With
            Next c
            GetDivHeight = accHeight + CalcFrameBorder(g, ascVertical) + 2 * MenuGrps(g).ContentsMarginV
    End Select

End Function

Public Function GetXYPos(g As Integer, IsForPreviewing As Boolean) As String

    If IsForPreviewing Or (Not Project.UserConfigs(Project.DefaultConfig).Frames.UseFrames) Then
        GetXYPos = CInt(MenuGrps(g).x) & ", " & CInt(MenuGrps(g).y)
    Else
        Select Case MenuGrps(g).Alignment
            Case gacBottomLeft, gacBottomRight, gacBottomCenter
                GetXYPos = "GetLeftTop()[0] + " & CInt(MenuGrps(g).x) & ", GetLeftTop()[1]"
            Case gacLeftBottom, gacLeftTop, gacLeftCenter
                GetXYPos = "GetWidthHeight()[0] - " & GetDivWidth(g) & ", GetLeftTop()[1] + " & CInt(MenuGrps(g).y)
            Case gacRightBottom, gacRightTop, gacRightCenter
                GetXYPos = "GetLeftTop()[0], GetLeftTop()[1] + " & CInt(MenuGrps(g).y)
            Case gacTopLeft, gacTopRight, gacTopCenter
                GetXYPos = "GetLeftTop()[0] + " & CInt(MenuGrps(g).x) & ", GetWidthHeight()[1] + GetLeftTop()[1] - " & GetDivHeight(g)
        End Select
    End If

End Function

Public Function GenFramesObject(ByVal sFrame As String) As String

    Dim Frames() As String
    Dim Obj As String
    Dim i As Integer
    
    If LenB(sFrame) = 0 Then sFrame = "_self"
    Select Case sFrame
        Case "_self", "_self."
            GenFramesObject = "_self"
        Case "_top", "_top."
            GenFramesObject = "top"
        Case "_blank", "_parent"
            GenFramesObject = sFrame
        Case Else
            If Right$(sFrame, 1) <> "." Then sFrame = sFrame + "."
            
            Frames = Split(sFrame, ".")
            For i = 0 To UBound(Frames) - 1
                Obj = Obj + "frames['" + Frames(i) + "']."
            Next i
            
            GenFramesObject = Left$(Obj, Len(Obj) - 1)
    End Select
    
    GenFramesObject = Replace$(GenFramesObject, "'", Chr(30))
    GenFramesObject = Replace$(GenFramesObject, """", Chr(29))

End Function

Public Function GetConfigID(ConfigName As String) As Integer

    Dim i As Integer
    
    For i = 0 To UBound(Project.UserConfigs)
        If Project.UserConfigs(i).Name = ConfigName Then
            GetConfigID = i
            Exit Function
        End If
    Next i
    
    GetConfigID = 0

End Function

Public Function ConvertPath(FileName As String) As String

    Dim RootWeb As String
    Dim tmpRootWeb As String
    Dim c As Integer
    Dim NewPath As String
    Dim ml As Integer
    
    NewPath = FileName
    For c = 0 To UBound(Project.UserConfigs)
        tmpRootWeb = Project.UserConfigs(c).RootWeb
        If tmpRootWeb = Left$(FileName, Len(tmpRootWeb)) And LenB(tmpRootWeb) <> 0 Then
            If Len(tmpRootWeb) > ml Then
                ml = Len(tmpRootWeb)
                RootWeb = tmpRootWeb
            End If
        End If
    Next c
    
    If RootWeb = Left$(FileName, Len(RootWeb)) And LenB(RootWeb) <> 0 Then
        NewPath = Project.UserConfigs(Project.DefaultConfig).RootWeb + Mid(FileName, Len(RootWeb) + 1)
    End If
    
    If InStr(NewPath, "/") Or InStr(NewPath, "\") Then
        Select Case Project.UserConfigs(Project.DefaultConfig).Type
            Case ctcRemote
                NewPath = SetSlashDir(NewPath, sdFwd)
            Case ctcLocal Or ctcCDROM
                If Not (UsesProtocol(NewPath) Or IsExternalLink(NewPath)) Then
                    NewPath = SetSlashDir(NewPath, sdBack)
                End If
        End Select
    End If
    
    ConvertPath = NewPath

End Function

Public Function px2pt(px As Integer) As Integer

    px2pt = CInt(px / (4 / 3))

End Function

Public Function pt2px(pt As Integer) As Integer

    pt2px = CInt(pt * (4 / 3))

End Function

Public Function GetParam(ByRef params As String, ByVal idx As Integer) As String

    On Error Resume Next
    If params <> "" Then GetParam = Split(params, cSep)(idx - 1)

End Function

Public Function GetHotSpotWidth(t As Integer, g As Integer, Optional IgnoreSettings As Boolean) As Integer
    
    Dim i As Integer
    Dim NoChilds As Boolean
    Dim cTB As ToolbarDef
    Dim oWidth As Integer
    Dim gSum As Integer
    Dim nHS As Integer
    
    Static acc As Integer
    Dim dw As Integer
    
    cTB = Project.Toolbars(t)
    
    NoChilds = True
    For i = 1 To UBound(MenuCmds)
        If MenuCmds(i).Parent = g Then
            NoChilds = False
            Exit For
        End If
    Next i
    
    With cTB
        oWidth = .Width
        If IgnoreSettings Then .Width = 0
        If (.Width = 0) Or (.Width = 1 And NoChilds) Then
            GetHotSpotWidth = GetTBItemWidth(g) + 2 * (MenuGrps(g).CmdsFXSize + MenuGrps(g).CmdsMarginX)
            .Width = oWidth
        Else
            Select Case .Width
                Case 1
                    GetHotSpotWidth = GetDivWidth(g) + CornerSize(g, "left") + CornerSize(g, "right")
                Case Else
                    If .JustifyHotSpots Or .Style = tscVertical Then
                        'r = GetTBWidth(t)
                        'GetHotSpotWidth = r(2)
                        GetHotSpotWidth = GetTBWidth(t)(2)
                    Else
                        If .Groups(1) = MenuGrps(g).Name Then acc = 0
                        For i = 1 To UBound(.Groups)
                            gSum = gSum + GetTBItemWidth(GetIDByName(.Groups(i)))
                            nHS = nHS + 1
                        Next i
                        If gSum > 0 Then
                            dw = GetTBWidth(t)(1)
                            GetHotSpotWidth = (dw - (2 * (.Border + .ContentsMarginH) + (nHS - 1) * .Separation)) * (GetTBItemWidth(g) / gSum)
                            acc = acc + GetHotSpotWidth
                            If cTB.Groups(UBound(.Groups)) = MenuGrps(g).Name Then
                                GetHotSpotWidth = GetHotSpotWidth + ((dw - 2 * (.Border + .ContentsMarginH) - (nHS - 1) * .Separation) - acc)
                            End If
                        End If
                    End If
            End Select
        End If
    End With
    
End Function

Public Function GetHotSpotHeight(t As Integer, g As Integer, Optional IgnoreSettings As Boolean) As Integer

    Dim cTB As ToolbarDef
    Dim oHeight As Integer
    Dim gSum As Integer
    Dim nHS As Integer
    Dim i As Integer
    
    Static acc As Integer
    Dim dH As Integer
    
    cTB = Project.Toolbars(t)
    
    With cTB
        oHeight = .Height
        If IgnoreSettings Then .Height = 0
        Select Case .Height
            Case 0
                GetHotSpotHeight = GetTBItemHeight(g) + 2 * (MenuGrps(g).CmdsFXSize + MenuGrps(g).CmdsMarginY)
                .Height = oHeight
            Case 1
                GetHotSpotHeight = GetDivWidth(g) + CornerSize(g, "top") + CornerSize(g, "bottom")
            Case Else
                If .JustifyHotSpots Or .Style = tscHorizonal Then
                    'r = GetTBHeight(t)
                    'GetHotSpotHeight = r(2)
                    GetHotSpotHeight = GetTBHeight(t)(2)
                Else
                    If .Groups(1) = MenuGrps(g).Name Then acc = 0
                    For i = 1 To UBound(.Groups)
                        gSum = gSum + GetTBItemHeight(GetIDByName(.Groups(i)))
                        nHS = nHS + 1
                    Next i
                    If gSum > 0 Then
                        dH = GetTBHeight(t)(1)
                        GetHotSpotHeight = (dH - (2 * (.Border + .ContentsMarginV) + (nHS - 1) * .Separation)) * (GetTBItemHeight(g) / gSum)
                        acc = acc + GetHotSpotHeight
                        If cTB.Groups(UBound(.Groups)) = MenuGrps(g).Name Then
                            GetHotSpotHeight = GetHotSpotHeight + ((dH - 2 * (.Border + .ContentsMarginH) - (nHS - 1) * .Separation) - acc)
                        End If
                    End If
                End If
        End Select
    End With
    
End Function

Private Function GetTBItemWidth(g As Integer) As Integer

    Dim tmpWidth As Integer

    With MenuGrps(g)
        If LenB(.Caption) <> 0 Then
            tmpWidth = GetGrpTextWidth(g)
        
            If LenB(.tbiLeftImage.NormalImage) <> 0 Then
                tmpWidth = tmpWidth + .tbiLeftImage.w + Preferences.ImgSpace
            End If
            If LenB(.tbiRightImage.NormalImage) <> 0 Then
                tmpWidth = tmpWidth + .tbiRightImage.w + 2 * Preferences.ImgSpace
            End If
            
            tmpWidth = tmpWidth + 2 * MenuGrps(g).CmdsMarginX
            
            applyCroppingWidth .tbiBackImage, tmpWidth
        Else
            If LenB(.tbiBackImage.NormalImage) <> 0 Then
                Set frmMain.picRsc.Picture = LoadPictureRes(.tbiBackImage.NormalImage)
                tmpWidth = frmMain.picRsc.Width / Screen.TwipsPerPixelX
            End If
            tmpWidth = .tbiLeftImage.w + tmpWidth + .tbiRightImage.w - 2 * .CmdsMarginX
        End If
    End With
    
    GetTBItemWidth = tmpWidth
    
End Function

Private Function GetTBItemHeight(g As Integer) As Integer

    Dim tmpHeight As Integer
    Dim ImgHeight As Integer

    With MenuGrps(g)
        If LenB(.Caption) <> 0 Then
            tmpHeight = GetGrpTextHeight(g)

            ImgHeight = IIf(.tbiLeftImage.h > .tbiRightImage.h, .tbiLeftImage.h, .tbiRightImage.h)
            If ImgHeight > tmpHeight Then tmpHeight = ImgHeight
            
            applyCroppingHeight .tbiBackImage, tmpHeight
        Else
            If LenB(.tbiBackImage.NormalImage) <> 0 Then
                Set frmMain.picRsc.Picture = LoadPictureRes(.tbiBackImage.NormalImage)
                tmpHeight = frmMain.picRsc.Height / Screen.TwipsPerPixelY
            End If
            'tmpHeight = IIf(.tbiLeftImage.h > .tbiRightImage.h, .tbiLeftImage.h, .tbiRightImage.h) + tmpHeight - 2 * .CmdsMarginY
            ImgHeight = IIf(.tbiLeftImage.h > .tbiRightImage.h, .tbiLeftImage.h, .tbiRightImage.h)
            ImgHeight = IIf(ImgHeight > tmpHeight, ImgHeight, tmpHeight)
            tmpHeight = ImgHeight - 2 * .CmdsMarginY
        End If
    End With

    GetTBItemHeight = tmpHeight

End Function

Public Function GetRGB(ByVal c As Long, Optional RetTransStr As Boolean = False) As String

    Dim sHex As String
    Dim r As String
    Dim g As String
    Dim b As String
    
    If c >= 0 Then
    
        sHex = Hex(c)
        sHex = String$(Abs(6 - Len(sHex)), "0") + sHex
        
        r = Right$(sHex, 2)
        g = Mid$(sHex, Len(sHex) - 3, 2)
        b = Mid$(sHex, Len(sHex) - 5, 2)
        
        GetRGB = "#" + r + g + b
    Else
        If c = -2 And RetTransStr Then
            GetRGB = "Transparent"
        Else
            GetRGB = c
        End If
    End If
    
End Function

Public Function GetRGBA(ByVal c As Long, a As Single) As String

    Dim sHex As String
    Dim r As String
    Dim g As String
    Dim b As String
    
    If c >= 0 Then
        sHex = Hex(c)
        sHex = String$(Abs(6 - Len(sHex)), "0") + sHex
        
        r = Right$(sHex, 2)
        g = Mid$(sHex, Len(sHex) - 3, 2)
        b = Mid$(sHex, Len(sHex) - 5, 2)
        
        GetRGBA = "rgba(" & Val("&h" + r) & "," & Val("&h" + g) & "," & Val("&h" + b) & "," & a & ")"
    Else
        GetRGBA = "rgba(0,0,0,0)"
    End If

End Function

Public Function GetAlignmentName(Alignment As AlignmentConstants) As String

    GetAlignmentName = IIf(Alignment = tacLeft, "left", IIf(Alignment = tacRight, "right", IIf(Alignment = tacCenter, "center", "left")))

End Function

Public Function CornerSize(g As Integer, Corner As String) As Integer
    'Corner can be one of these: Left, Top, Right or Bottom
    
    Dim c1 As Integer
    Dim c2 As Integer
    Dim c3 As Integer
    Dim mgc As GroupCorners
   
    mgc = MenuGrps(g).CornersImages
    
    Select Case LCase(Corner)
        Case "left"
            c1 = GetImageSize(mgc.gcTopLeft)(0)
            c2 = GetImageSize(mgc.gcLeft)(0)
            c3 = GetImageSize(mgc.gcBottomLeft)(0)
        Case "top"
            c1 = GetImageSize(mgc.gcTopLeft)(1)
            c2 = GetImageSize(mgc.gcTopCenter)(1)
            c3 = GetImageSize(mgc.gcTopRight)(1)
        Case "right"
            c1 = GetImageSize(mgc.gcTopRight)(0)
            c2 = GetImageSize(mgc.gcRight)(0)
            c3 = GetImageSize(mgc.gcBottomRight)(0)
        Case "bottom"
            c1 = GetImageSize(mgc.gcBottomLeft)(1)
            c2 = GetImageSize(mgc.gcBottomCenter)(1)
            c3 = GetImageSize(mgc.gcBottomRight)(1)
    End Select
        
    CornerSize = IIf(c1 > c2, IIf(c1 > c3, c1, c3), IIf(c2 > c3, c2, c3))
    
End Function

Public Function GetImageSize(FileName As String) As Integer()

    Dim s(0 To 1) As Integer
    
    If LenB(FileName) = 0 Then
        s(0) = 0
        s(1) = 0
    Else
        With frmMain.picRsc
            .Picture = LoadPictureRes(FileName)
            s(0) = .Width / Screen.TwipsPerPixelX
            s(1) = .Height / Screen.TwipsPerPixelY
        End With
    End If
    
    GetImageSize = s

End Function

Public Function MemberOf(g As Integer) As Integer

    Dim t As Integer
    Dim gs As String
    
    For t = 1 To UBound(Project.Toolbars)
        With Project.Toolbars(t)
            gs = Join(.Groups, "|")
            If InStr(gs, MenuGrps(g).Name) > 0 Then
                MemberOf = t
                Exit Function
            End If
        End With
    Next t

End Function

Public Function MemberOfByName(gName As String) As Integer

    Dim t As Integer
    Dim gs As String
    
    For t = 1 To UBound(Project.Toolbars)
        With Project.Toolbars(t)
            gs = Join(.Groups, "|")
            If InStr(gs, gName) > 0 Then
                MemberOfByName = t
                Exit Function
            End If
        End With
    Next t

End Function

Public Function CreateToolbar() As Boolean

    On Error Resume Next
    
    Dim i As Integer

    CreateToolbar = (UBound(Project.Toolbars) > 0)
    If CreateToolbar Then
        For i = 1 To UBound(Project.Toolbars)
            CreateToolbar = (UBound(Project.Toolbars(i).Groups) > 0) Or CreateToolbar
        Next i
    End If

End Function

Public Function ToolbarIndexByKey(TBKey As String) As Integer

    ToolbarIndexByKey = ToolbarIndexByName(Mid(TBKey, 4))

End Function

Public Function ToolbarIndexByName(TBName As String) As Integer

    Dim i As Integer

    For i = 1 To UBound(Project.Toolbars)
        If Project.Toolbars(i).Name = TBName Then
            ToolbarIndexByName = i
            Exit Function
        End If
    Next i

End Function

Public Function EmptyToolbars() As Boolean

    Dim i As Integer
    
    EmptyToolbars = True
    For i = 1 To UBound(Project.Toolbars)
        EmptyToolbars = EmptyToolbars And (UBound(Project.Toolbars(i).Groups) = 0)
    Next i

End Function

Public Function GetRealSubGroup(ByVal g As Integer, Optional IgnoreIfNull As Boolean = False) As Integer

    Dim k As Integer
    
    With MenuGrps(g).Actions
        If .OnMouseOver.Type = atcCascade Then
            k = .OnMouseOver.TargetMenu
        End If
        If .OnClick.Type = atcCascade Then
            k = .OnClick.TargetMenu
        End If
        If .OnDoubleClick.Type = atcCascade Then
            k = .OnDoubleClick.TargetMenu
        End If
    End With

    If IgnoreIfNull Then
        GetRealSubGroup = k
    Else
        GetRealSubGroup = IIf(k = 0, g, k)
    End If

End Function

Public Function NiceCmdCaption(ItemID As Integer) As String

    Dim ItemCaption As String

    ItemCaption = xUNI2Unicode(RemoveHTMLCode(MenuCmds(ItemID).Caption))
    
    If LenB(ItemCaption) = 0 Then
        ItemCaption = "[" + MenuCmds(ItemID).Name + "]"
    End If
    
    NiceCmdCaption = ItemCaption

End Function

Public Function NiceGrpCaption(ItemID As Integer) As String

    Dim ItemCaption As String

    ItemCaption = xUNI2Unicode(RemoveHTMLCode(MenuGrps(ItemID).Caption))
    
    If LenB(ItemCaption) = 0 Then
        ItemCaption = "[" + MenuGrps(ItemID).Name + "]"
    End If
    
    NiceGrpCaption = ItemCaption

End Function

'Public Function DecodePwd(pwd As String) As String
'
'    Dim p As String
'    Dim t As String
'    Dim i As Integer
'
'    p = Left(pwd, 32) + Space(32 - Len(pwd))
'    For i = 1 To Len(p) / 2 Step 2
'        t = Mid(p, i, 1)
'        Mid(p, i, 1) = Mid(p, Len(p) - i + 1, 1)
'        Mid(p, Len(p) - i + 1, 1) = t
'    Next i
'
'    XFX_Decode p, 32
'
'    DecodePwd = Trim(p)
'
'End Function

'Public Function EncodePwd(pwd As String) As String
'
'    Dim p As String
'    Dim t As String
'    Dim i As Integer
'
'    p = pwd + Space(32 - Len(pwd))
'    XFX_Encode p
'
'    For i = 1 To Len(p) / 2 Step 2
'        t = Mid(p, i, 1)
'        Mid(p, i, 1) = Mid(p, Len(p) - i + 1, 1)
'        Mid(p, Len(p) - i + 1, 1) = t
'    Next i
'
'    EncodePwd = p
'
'End Function

Public Sub MkDir2(ByVal d As String)

          Dim i As Integer
          Dim sd() As String
          Dim ss As String
          
    On Error GoTo MkDir2_Error

10        sd = Split(d, "\")
20        For i = 0 To UBound(sd)
30            ss = Replace(ss + sd(i) + "\", "\\", "\")
40            If i > 0 Then
50                If Not FolderExists(ss) Then MkDir ss
60            End If
70        Next i

    On Error GoTo 0
    Exit Sub

MkDir2_Error:

    MsgBox "Error " & Err.Number & " in line " & Erl & ": " & vbCrLf & Err.Description & " in Module.modGlobals.MkDir2"
    Err.Raise Err.Number

End Sub
