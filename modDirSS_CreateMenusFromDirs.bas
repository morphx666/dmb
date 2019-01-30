Attribute VB_Name = "modDirSS_CreateMenusFromDirs"
Option Explicit

Public DirSSValidFileTypes As String
Private DefGrp As String

Public Enum dirssTextMode
    [tmDocTitle]
    [tmFileName]
    [tmLinkText]
End Enum
Private xTextMode As dirssTextMode

Private Type LinkInfo
    href As String
    Text As String
End Type

Private FloodPanel As New clsFlood
Private numItems As Long
Private cItem As Long

Public Sub StartLinksScan(ByVal fName As String, TextMode As dirssTextMode)

    xTextMode = TextMode
    DefGrp = GetGrpParams(TemplateGroup)
    
    Set FloodPanel.PictureControl = frmDirSS.picFlood
    
    AddLinks2Toolbar GetLinks(fName)

End Sub

Private Function GetLinks(fName As String) As LinkInfo()

    Dim sCode As String
    Dim a() As String
    Dim l() As LinkInfo
    Dim i As Long
    Dim j As Long
    Dim href As String
    Dim Exists As Boolean
    Dim txt As String
    
    sCode = LoadFile(fName)
    a = Split(LCase(sCode), "<a ")
    
    ReDim l(0)
    
    For i = 1 To UBound(a)
        href = GetParamVal(a(i), "href")
        If Not (UsesProtocol(href) Or IsExternalLink(href)) Then
            href = SetSlashDir(href, sdBack)
            If Left(href, 8) = "file:\\\" Then
                href = Replace(Mid(href, 9), "|", ":")
            End If
            If (Left(href, 1) >= 0 And Left(href, 1) <= 9) Or (Left(href, 1) >= "a" And Left(href, 1) <= "z") Or Left(href, 1) = "_" Or Left(href, 1) = "." Or Left(href, 1) = "/" Or Left(href, 1) = "\" Then
                href = Replace(GetFilePath(fName) + href, "\\", "\")
            End If
            
            If FileExists(href) Then
                Exists = False
                For j = 1 To UBound(l)
                    If l(j).href = href Then
                        Exists = True
                        Exit For
                    End If
                Next j
                If Not Exists Then
                    txt = GetLinkText(href, a(i), sCode)
                    If LenB(txt) <> 0 Then
                        ReDim Preserve l(UBound(l) + 1)
                        With l(UBound(l))
                            .href = Short2Long(Long2Short(href))
                            .Text = txt
                        End With
                    End If
                End If
            End If
        End If
    Next i
    
    GetLinks = l

End Function

Private Function GetLinkText(href As String, atag As String, sCode As String) As String

    Dim txt As String
    Dim p As Long

    Select Case xTextMode
        Case tmDocTitle
            txt = GetDocTITLE(href)
        Case tmFileName
            txt = GetFileName(href, True)
        Case tmLinkText
            txt = "<a " + Left(atag, InStr(atag, "</a>") + 3)
            p = InStr(LCase(sCode), txt)
            txt = Mid(sCode, p, Len(txt))
            txt = RemoveHTMLCode(txt)
            If LenB(txt) = 0 Then txt = GetDocTITLE(href)
            If LenB(txt) = 0 Then txt = GetFileName(href, True)
    End Select
    txt = Replace(txt, vbCrLf, "<br>")
    txt = Replace(txt, vbTab, "")
    If txt = "<br>" Then txt = ""
    
    GetLinkText = txt

End Function

Private Sub AddLinks2Toolbar(Links() As LinkInfo)

    Dim g As Integer
    Dim i As Integer
    
    For i = 1 To UBound(Links)
        g = AddGroupFromLink(Links(i), False)
        
        If g > 0 Then
            ReDim Preserve Project.Toolbars(1).Groups(UBound(Project.Toolbars(1).Groups) + 1)
            Project.Toolbars(1).Groups(UBound(Project.Toolbars(1).Groups)) = MenuGrps(g).Name
        End If
    Next i
    
    For i = 1 To UBound(Project.Toolbars(1).Groups)
        g = GetIDByName(Project.Toolbars(1).Groups(i))
        If AddCommands2Group(g, Links(i).href) = 0 Then
            MenuGrps(g).Actions.onmouseover.Type = atcNone
        End If
    Next i

End Sub

Private Function AddGroupFromLink(l As LinkInfo, Optional ByVal Ok2AddCommands As Boolean = True) As Integer

    Dim g As Integer
    
    For g = 1 To UBound(MenuGrps)
        If MenuGrps(g).Actions.onclick.url = l.href Then Exit Function
    Next g

    AddMenuGroup DefGrp, True
    g = UBound(MenuGrps)
    With MenuGrps(g).Actions.onmouseover
        .Type = atcCascade
        .TargetMenu = g
    End With
    
    DoEvents

    With MenuGrps(g)
        .Caption = l.Text
        .Actions.onclick.Type = atcURL
        .Actions.onclick.url = l.href
    End With
    
    If Ok2AddCommands Then
        If AddCommands2Group(g, l.href) = 0 Then
            ReDim Preserve MenuGrps(UBound(MenuGrps) - 1)
            g = 0
        End If
    End If
    
    AddGroupFromLink = g

End Function

Private Function AddCommands2Group(ByVal g As Integer, ByVal fName As String) As Integer

    Dim l() As LinkInfo
    Dim i As Integer
    Dim c As Integer
    Dim Exists As Boolean
    Dim k As Integer
    
    Static recursion As Integer
    
    recursion = recursion + 1
    
    l = GetLinks(fName)
    
    For i = 1 To UBound(l)
        Exists = False
        For c = 1 To UBound(MenuCmds)
            If MenuCmds(c).Actions.onclick.url = l(i).href Then
                Exists = True
                Exit For
            End If
        Next c
        If Not Exists Then
            k = k + 1
            With TemplateCommand
                .Caption = l(i).Text
                .Actions.onclick.Type = atcURL
                .Actions.onclick.url = l(i).href
                .Parent = g
            End With
            AddMenuCommand GetCmdParams(TemplateCommand), True, True, True
            
            numItems = numItems + 1
            If numItems > 30 Then numItems = 0
            
            FloodPanel.Caption = "Scanning: '" + EllipseText(frmDirSS.txtDummy, fName, DT_PATH_ELLIPSIS) + "'"
            FloodPanel.Value = numItems / 30 * 100
            
            c = UBound(MenuCmds)
            If recursion < 10 Then
                If UBound(GetLinks(l(i).href)) > 0 Then
                    MenuCmds(c).Actions.onmouseover.Type = atcCascade
                    MenuCmds(c).Actions.onmouseover.TargetMenuAlignment = gacRightTop
                    MenuCmds(c).Actions.onmouseover.TargetMenu = AddGroupFromLink(l(i))
                    If MenuCmds(c).Actions.onmouseover.TargetMenu = 0 Then
                        MenuCmds(c).Actions.onmouseover.Type = atcNone
                    Else
                        MenuCmds(c).Caption = MenuCmds(c).Caption + " »"
                    End If
                End If
            End If
        End If
    Next i
    
    recursion = recursion - 1
    
    AddCommands2Group = k

End Function

'-------------------------------------
'-------------------------------------
'-------------------------------------

Public Function StartDirScan(ByVal fName As String, TextMode As dirssTextMode, GroupRootDocs As Boolean)

    xTextMode = TextMode
    DefGrp = GetGrpParams(TemplateGroup)
    
    Set FloodPanel.PictureControl = frmDirSS.picFlood
    
    'CreateGroupFromFolder fName
    CreateToolbarFromFolder fName, GroupRootDocs

End Function

Private Sub CreateToolbarFromFolder(ByVal fName As String, GroupRootDocs As Boolean)

    Dim dFile As String
    Dim fldr() As String
    Dim File() As String
    Dim ng As Integer
    Dim i As Integer
    
    ReDim fldr(0)
    ReDim File(0)
    
    dFile = Dir(fName, vbDirectory)
    Do While LenB(dFile) <> 0
        If dFile <> "." And dFile <> ".." Then
            If ((GetAttr(fName + dFile) And vbDirectory) = vbDirectory) Then
                ReDim Preserve fldr(UBound(fldr) + 1)
                fldr(UBound(fldr)) = dFile
            Else
                If MatchSpec(dFile, DirSSValidFileTypes) Then
                    If Left(dFile, 5) <> "_vti_" Then
                        ReDim Preserve File(UBound(File) + 1)
                        File(UBound(File)) = dFile
                    End If
                End If
            End If
        End If
        dFile = Dir
    Loop
    
    For i = 1 To UBound(fldr)
        If Not IsFrontPageSysDir(fldr(i)) Then
            AddGroup2ToolbarFromFolder fName + fldr(i), False
        End If
    Next i
    
    If GroupRootDocs Then
        If UBound(File) > 0 Then
            ng = AddGroup2ToolbarFromFolder(fName, False, False)
        End If
    Else
        For i = 1 To UBound(File)
            FloodPanel.Caption = "Scanning: " + fName + File(i)
            AddGroup2ToolbarFromFolder fName + File(i), True, False
        Next i
    End If
    
End Sub

Private Function AddGroup2ToolbarFromFolder(fName As String, IsFile As Boolean, Optional IncludeFolders As Boolean = True) As Integer

    Dim ng As Integer

    ng = CreateGroupFromFolder(fName, IncludeFolders, Not IsFile)
    If ng > 0 Then
        ReDim Preserve Project.Toolbars(1).Groups(UBound(Project.Toolbars(1).Groups) + 1)
        Project.Toolbars(1).Groups(UBound(Project.Toolbars(1).Groups)) = MenuGrps(ng).Name
    End If
    
    AddGroup2ToolbarFromFolder = ng

End Function

Private Function CreateGroupFromFolder(ByVal fName As String, Optional IncludeFolders As Boolean = True, Optional IsFolder As Boolean = True) As Integer

    Dim g As Integer
    
    If IsFolder Then fName = AddTrailingSlash(fName, "\")
    
    AddMenuGroup DefGrp, True
    g = UBound(MenuGrps)
    With MenuGrps(g).Actions.onmouseover
        .Type = atcCascade
        .TargetMenu = g
    End With
    
    If IsFolder Then
        FloodPanel.Caption = "Scanning: '" + EllipseText(frmDirSS.txtDummy, fName, DT_PATH_ELLIPSIS) + "'"
        MenuGrps(g).Caption = GetFolderName(fName)
        If AddCommandsFromFiles(fName, g, IncludeFolders) > 0 Then
            CreateGroupFromFolder = g
        Else
            ReDim Preserve MenuGrps(g - 1)
            CreateGroupFromFolder = 0
        End If
    Else
        Select Case xTextMode
            Case tmDocTitle
                MenuGrps(g).Caption = GetDocTITLE(fName)
            Case tmFileName
                MenuGrps(g).Caption = GetFileName(fName, True)
        End Select
        MenuGrps(g).Actions.onmouseover.Type = atcNone
        With MenuGrps(g).Actions.onclick
            .Type = atcURL
            .url = fName
        End With
        frmMain.tvMenus.Nodes("G" & UBound(MenuGrps)).Image = GenGrpIcon(UBound(MenuGrps))
        CreateGroupFromFolder = g
    End If
    
End Function

Private Function GetFolderName(ByVal fName As String)

    Dim gc() As String
    
    gc = Split(fName, "\")
    GetFolderName = gc(UBound(gc) - 1)

End Function

Private Function IsFrontPageSysDir(ByVal fName As String) As Boolean

    IsFrontPageSysDir = fName = "_notes" Or fName = "_borders" Or fName = "_colab" Or Left(fName, 5) = "_vti_" Or fName = "_derived" Or fName = "_overlay" Or fName = "_private" Or fName = "_borders"

End Function

Private Function AddCommandsFromFiles(ByVal fName As String, ByVal g As Integer, Optional IncludeFolders As Boolean = True) As Integer

    Dim dFile As String
    Dim fldr() As String
    Dim File() As String
    Dim i As Integer
    Dim ng As Integer
    Dim k As Integer
    
    ReDim fldr(0)
    ReDim File(0)
    
    DoEvents
    
    dFile = Dir(fName, vbDirectory)
    Do While LenB(dFile) <> 0
        If dFile <> "." And dFile <> ".." Then
            If ((GetAttr(fName + dFile) And vbDirectory) = vbDirectory) Then
                If Not IsFrontPageSysDir(dFile) Then
                    ReDim Preserve fldr(UBound(fldr) + 1)
                    fldr(UBound(fldr)) = dFile
                End If
            Else
                If MatchSpec(dFile, DirSSValidFileTypes) Then
                    ReDim Preserve File(UBound(File) + 1)
                    File(UBound(File)) = dFile
                End If
            End If
        End If
        dFile = Dir
    Loop
    
    If IncludeFolders Then
        numItems = numItems + UBound(fldr)
        For i = 1 To UBound(fldr)
            cItem = cItem + 1
            FloodPanel.Value = cItem / numItems * 100
            ng = CreateGroupFromFolder(fName + fldr(i))
            If ng > 0 Then
                With TemplateCommand
                    .Caption = fldr(i) + " »"
                    .Actions.onclick.Type = atcNone
                    .Actions.onmouseover.Type = atcCascade
                    .Actions.onmouseover.TargetMenu = ng
                    .Actions.onmouseover.TargetMenuAlignment = gacRightTop
                    .Parent = g
                End With
                AddMenuCommand GetCmdParams(TemplateCommand), True, True, True
                k = k + 1
            End If
        Next i
    End If
    
    numItems = numItems + UBound(File)
    For i = 1 To UBound(File)
        cItem = cItem + 1
        FloodPanel.Value = cItem / numItems * 100
        With TemplateCommand
            Select Case xTextMode
                Case tmDocTitle
                    .Caption = GetDocTITLE(fName + File(i))
                    If LenB(.Caption) = 0 Then
                        If Not MatchSpec(File(i), "*." + Join(Split(strGetSupportedHTMLDocs, ";"), ";*.")) Then
                            .Caption = GetFileName(File(i), True)
                        End If
                    End If
                Case tmFileName
                    .Caption = GetFileName(File(i), True)
            End Select
            
            If LenB(.Caption) <> 0 Then
                .Actions.onclick.Type = atcURL
                .Actions.onclick.url = fName + File(i)
                .Actions.onmouseover.Type = atcNone
                .Parent = g

                AddMenuCommand GetCmdParams(TemplateCommand), True, True, True
                k = k + 1
            End If
        End With
    Next i
    
    AddCommandsFromFiles = k

End Function

Private Function GetDocTITLE(ByVal mFileName As String) As String

    Dim sCode As String
    Dim p1 As Long
    Dim p2 As Long
    
    On Error Resume Next
    
    sCode = LoadFile(mFileName)
    
    p1 = InStr(1, sCode, "<title>", vbTextCompare)
    p2 = InStr(p1, sCode, "</title>", vbTextCompare)
    If p2 > p1 Then
        sCode = Mid(sCode, p1 + 7, p2 - p1 - 7)
        sCode = Replace(sCode, vbCrLf, "")
        sCode = Replace(sCode, vbCr, "")
        sCode = Replace(sCode, vbLf, "")
    Else
        sCode = ""
    End If
    
    GetDocTITLE = sCode

End Function
