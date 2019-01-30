Attribute VB_Name = "modImportEngine"
Option Explicit

Public Sub ImportProject(FileName As String)

    Dim g As Integer
    On Error GoTo ReportError

    Select Case AnalyzeImportProject(FileName)
        Case "hierMenus"
            Import_hierMenus FileName
        Case "hierMenus4"
            Import_hierMenus4 FileName
        Case "AllWebMenus"
            Import_AWM FileName
        Case "ApPopupMenu"
            Import_ApPopUpMenu FileName
        Case Else
            MsgBox "The selected document is not a recognized project"
    End Select
    
    If Project.ToolBar.CreateToolbar Then
        Project.Toolbars(1) = Project.ToolBar
        ReDim Project.Toolbars(1).Groups(0)
        For g = 1 To UBound(MenuGrps)
            If MenuGrps(g).IncludeInToolbar Then
                ReDim Preserve Project.Toolbars(1).Groups(UBound(Project.Toolbars(1).Groups) + 1)
                Project.Toolbars(1).Groups(UBound(Project.Toolbars(1).Groups)) = MenuGrps(g).Name
            End If
        Next g
    End If
    
    Exit Sub
    
ReportError:
    MsgBox "The Import process has failed", vbInformation + vbOKOnly, "Error Importing"

End Sub

Private Function AnalyzeImportProject(FileName As String) As String

    Dim sCode As String

    sCode = LoadFile(FileName)
    
    If InStr(1, sCode, "new array(", vbTextCompare) Then
        AnalyzeImportProject = "hierMenus"
        Exit Function
    End If
    
    If InStr(1, sCode, "HM_Array1", vbTextCompare) Then
        AnalyzeImportProject = "hierMenus4"
        Exit Function
    End If
    
    If Left(sCode, 3) = "awm" Or Left(sCode, 2) = "M0" Then
        AnalyzeImportProject = "AllWebMenus"
        Exit Function
    End If
    
    If InStr(sCode, "{|") Then
        AnalyzeImportProject = "ApPopupMenu"
        Exit Function
    End If

End Function

Private Sub Import_AWM(FileName As String)

    Dim sCode As String
    Dim lines() As String
    Dim l As Integer
    Dim mPar() As String
    Dim NewCmd As MenuCmd
    Dim NewGrp As MenuGrp
    Dim gName As String
    Dim Is13 As Boolean
    Dim j As Integer
    Dim cName As String
    
    sCode = LoadFile(FileName)

    lines = Split(sCode, vbCrLf)
    
    Is13 = Val(Mid(lines(0), 4)) >= 130
    If Not Is13 Then
        For l = 0 To UBound(lines) - 1
            If UBound(Split(lines(l), ",")) > 0 Then
                If Left(lines(l - 1), 1) <> "I" Then
                    mPar = Split(lines(l), ",")
                    ReDim Preserve mPar(UBound(mPar) + 1)
                    For j = UBound(mPar) - 1 To 1 Step -1
                        mPar(j) = mPar(j - 1)
                    Next j
                    lines(l) = Join(mPar, ",")
                End If
            End If
        Next l
    End If
    
    For l = 0 To UBound(lines) - 1
        If Left(lines(l), 2) = "M0" Then
            Exit For
        End If
    Next l
    
    For l = l To UBound(lines) - 1 Step 4
        If Left(lines(l), 1) = "M" Or Left(lines(l), 1) = "S" Then
            mPar = Split(lines(l + 1), ",")
            mPar(0) = Replace(mPar(0), "!", "")
            NewGrp = TemplateGroup
            With NewGrp
                .Name = RemoveRepNames(mPar(0) + "_" + lines(l), 1)

                If Left(lines(l), 1) = "S" Then
                    .Caption = Replace(Split(lines(l - 3), ",")(0), "!", "")
                    .DefNormalFont.FontName = mPar(6)
                    .DefNormalFont.FontSize = pt2px(Val(mPar(7)) + 1) * 3
                    .DefNormalFont.FontBold = -Val(mPar(8))
                    .DefNormalFont.FontItalic = -Val(mPar(9))
                    .DefNormalFont.FontUnderline = -Val(mPar(10))
                    If mPar(11) <> "n" Then .nTextColor = HexColor2OLEColor(mPar(11))
                    If mPar(16) <> "n" Then .nBackColor = HexColor2OLEColor(mPar(16))
                    .IncludeInToolbar = (Left(lines(l - 6), 1) = "M")
                Else
                    .Caption = Replace(mPar(0), "!", "")
                    .DefNormalFont.FontName = mPar(14)
                    .DefNormalFont.FontSize = pt2px(Val(mPar(15)) + 1) * 3
                    .DefNormalFont.FontBold = -Val(mPar(16))
                    .DefNormalFont.FontItalic = -Val(mPar(17))
                    .DefNormalFont.FontUnderline = -Val(mPar(18))
                    If mPar(19) <> "n" Then .nTextColor = HexColor2OLEColor(mPar(19))
                    If mPar(24) <> "n" Then .nBackColor = HexColor2OLEColor(mPar(24))
                    .IncludeInToolbar = True
                End If
                .WinStatus = .Caption
                .DefHoverFont = .DefNormalFont
                .ToolbarIndex = UBound(MenuGrps) + 1
                                
                .Actions.onmouseover.Type = atcCascade
                .Actions.onmouseover.TargetMenu = UBound(MenuGrps)
            End With
            AddMenuGroup GetGrpParams(NewGrp)
        Else
            mPar = Split(lines(l + 1), ",")
            cName = RemoveRepNames(Trim(Replace(mPar(0), "!", "")), 0)
            If cName <> "" Then
                NewCmd = TemplateCommand
                With NewCmd
                    .Name = cName
                    .Caption = Trim(Replace(mPar(0), "!", ""))
                    .WinStatus = Replace(mPar(51), "!", "")
                    .Parent = awbFindParent(lines, lines(l - 2))
                    
                    If .Parent > 0 Then
                        .NormalFont.FontName = mPar(15)
                        .HoverFont.FontName = mPar(16)
                        .NormalFont.FontSize = pt2px(Val(mPar(18)) + 1) * 3
                        .HoverFont.FontSize = pt2px(Val(mPar(19)) + 1) * 3
                        .NormalFont.FontBold = -Val(mPar(21))
                        .HoverFont.FontBold = -Val(mPar(22))
                        .NormalFont.FontItalic = -Val(mPar(24))
                        .HoverFont.FontItalic = -Val(mPar(25))
                        .NormalFont.FontUnderline = -Val(mPar(27))
                        .HoverFont.FontUnderline = -Val(mPar(28))
                        
                        .nTextColor = HexColor2OLEColor(mPar(30))
                        .hTextColor = HexColor2OLEColor(mPar(31))
                        
                        If mPar(45) <> "n" Then .nBackColor = HexColor2OLEColor(mPar(45))
                        If mPar(46) <> "n" Then .hBackColor = HexColor2OLEColor(mPar(46))
                        
                        .Actions.onclick.Type = atcURL
                        .Actions.onclick.url = IIf(mPar(53) <> "n", mPar(53), "")
                        
                        If lines(l + 2) = lines(l) Then
                            If lines(l + 3) = Split(lines(l + 5), ",")(0) Then
                                .Actions.onmouseover.Type = atcCascade
                                .Actions.onmouseover.TargetMenu = .Parent + 1
                            End If
                        End If
                        
                        frmMain.tvMenus.Nodes("G" & .Parent).Selected = True
                    End If
                End With
                AddMenuCommand GetCmdParams(NewCmd)
            End If
        End If
    Next l
    
End Sub

Private Function RemoveRepNames(ByVal n As String, ItemType As Integer) As String

    Dim i As Integer
    Dim cn As String
    Dim tn As String
    Dim K As Integer
    
    n = FixItemName(Trim(n))
    If n = "" Then Exit Function

    If ItemType = 0 Then
        K = 0
        tn = n
ReStartCmd:
        For i = 1 To UBound(MenuCmds)
            cn = MenuCmds(i).Name
            If tn = cn Then
                tn = n + "_" + Format(K, "000")
                K = K + 1
                GoTo ReStartCmd
            End If
        Next i
    End If
    
    If ItemType = 1 Then
        K = 0
        tn = n
ReStartGrp:
        For i = 1 To UBound(MenuGrps)
            cn = MenuGrps(i).Name
            If tn = cn Then
                tn = n + "_" + Format(K, "000")
                K = K + 1
                GoTo ReStartGrp
            End If
        Next i
    End If
    
    RemoveRepNames = IIf(tn = "", n, tn)

End Function

Private Function awbFindParent(lines() As String, pName As String) As Integer

    Dim l As Integer
    Dim tname As String

    For l = 1 To UBound(lines) - 1
        If lines(l) = pName Then
            tname = FixItemName(Replace(Replace(Split(lines(l + 1), ",")(0) + "_" + pName, " ", "_"), "!", ""))
            If Left(tname, 1) = "_" Then tname = Mid(tname, 2)
            awbFindParent = GetIDByName(tname)
            Exit For
        End If
    Next l

End Function

Private Function HexColor2OLEColor(ByVal h As String) As Long

    h = Mid(h, 2)
    HexColor2OLEColor = RGB(Hex2Dec(Left(h, 2)), Hex2Dec(Mid(h, 3, 2)), Hex2Dec(Right(h, 2)))

End Function

Private Function Hex2Dec(h As String) As Long

    On Error Resume Next

    If h = "" Then h = 0
    Hex2Dec = CLng("&H" & h)

End Function

Private Sub Import_hierMenus(FileName As String)

    Dim sCode As String
    Dim lines() As String
    Dim l As Integer
    Dim par As String
    Dim menus() As String
    Dim IsV3 As Boolean
    Dim IsSubMenu As Boolean
    Dim mPar() As String
    Dim p As Integer
    Dim nPar As String
    Dim NewCmd As MenuCmd
    Dim NewGrp As MenuGrp
    
    sCode = LoadFile(FileName)
    
    lines = Split(sCode, vbLf)
    ReDim menus(0)
    For l = 0 To UBound(lines)
        If InStr(LCase(lines(l)), "new array") Then
            If InStr(par, ")") Then
                par = Left(par, InStr(par, ")") - 1)
                ReDim Preserve menus(UBound(menus) + 1)
                menus(UBound(menus)) = Replace(par, """", "")
                par = Trim(Split(lines(l), "=")(0)) + ","
            Else
                par = Trim(Split(lines(l), "=")(0)) + ","
            End If
        Else
            par = par + lines(l)
        End If
    Next l
    
    IsV3 = MsgBox("Is this array for hierMenus 3.0 or above?", vbQuestion + vbYesNo, "hierMenus Version") = vbYes
    
    For l = 1 To UBound(menus)
        mPar = Split(Split(menus(l), cSep)(0), ",")
        IsSubMenu = (InStr(mPar(0), "_") > 0)
        
        NewGrp = TemplateGroup
        With NewGrp
            .Name = mPar(0)
            .Caption = .Name
            .Actions.onmouseover.Type = atcCascade
            .Actions.onmouseover.TargetMenu = UBound(MenuGrps)
            .ToolbarIndex = UBound(MenuGrps) + 1
        End With
        AddMenuGroup GetGrpParams(NewGrp)
        
        If IsV3 And Not IsSubMenu Then
            For p = 10 To UBound(mPar) Step 3
                NewCmd = TemplateCommand
                With NewCmd
                    .Name = mPar(p)
                    .Caption = mPar(p)
                    .WinStatus = mPar(p)
                    .Actions.onclick.Type = atcURL
                    .Actions.onclick.url = mPar(p + 1)
                    If Val(mPar(p + 2)) = 1 Then
                        .Actions.onmouseover.Type = atcCascade
                        .Actions.onmouseover.TargetMenu = UBound(MenuGrps) + 1
                    End If
                End With
                AddMenuCommand GetCmdParams(NewCmd)
            Next p
        Else
            For p = 1 To UBound(mPar) Step 3
                NewCmd = TemplateCommand
                With NewCmd
                    .Name = mPar(p)
                    .Caption = mPar(p)
                    .WinStatus = mPar(p)
                    .Actions.onclick.Type = atcURL
                    .Actions.onclick.url = mPar(p + 1)
                    If Val(mPar(p + 2)) = 1 Then
                        .Actions.onmouseover.Type = atcCascade
                        .Actions.onmouseover.TargetMenu = UBound(MenuGrps) + 1
                    End If
                End With
                AddMenuCommand GetCmdParams(NewCmd)
            Next p
        End If
    Next l

End Sub

Private Sub Import_ApPopUpMenu(FileName As String)

    On Error Resume Next
    
    Dim lines() As String
    Dim sCode As String
    Dim c As Integer
    Dim g As Integer
    Dim HasCmds As Boolean
    
    Project.ToolBar.CreateToolbar = True
    
    sCode = LoadFile(FileName)
    lines = Split(sCode, vbLf)
    
    FloodPanel.Caption = "Importing..."
    
    Import_ApPopUpMenu_AddGroup lines, 0
    
    For g = 1 To UBound(MenuGrps)
        HasCmds = False
        For c = 1 To UBound(MenuCmds)
            If MenuCmds(c).Parent = g Then
                HasCmds = True
                Exit For
            End If
        Next c
        If Not HasCmds Then MenuGrps(g).IncludeInToolbar = False
    Next g
    
    FloodPanel.Value = 0

End Sub

Private Function Import_ApPopUpMenu_AddGroup(lines() As String, l As Integer) As Integer

    Dim NewCmd As MenuCmd
    Dim NewGrp As MenuGrp
    Static level As Integer
    Dim p() As String
    Dim nl As Integer
    
    If l = 0 Then level = 0
    
    level = Import_ApPopUpMenu_GetLevel(lines(l))
    p = Import_ApPopUpMenu_pars(lines(l), level)
    
    NewGrp = TemplateGroup
    NewGrp.Caption = p(0)
    NewGrp.Name = FixItemName(p(0)) + Format(UBound(MenuGrps), "0000")
    NewGrp.IncludeInToolbar = (level = 0)
    NewGrp.ToolbarIndex = UBound(MenuGrps) + 1
    AddMenuGroup GetGrpParams(NewGrp)
    
    DoEvents

    For l = l + 1 To UBound(lines)
        FloodPanel.Value = (l / UBound(lines)) * 100
        nl = Import_ApPopUpMenu_GetLevel(lines(l))
        If nl = level + 1 Then
            p = Import_ApPopUpMenu_pars(lines(l), level)

            NewCmd = TemplateCommand
            p(0) = Mid(p(0), 2)
            If p(0) = "-" Then
                NewCmd.Name = "[SEP]"
            Else
                NewCmd.Caption = p(0)
                NewCmd.Name = FixItemName(NewCmd.Caption) + Format(UBound(MenuCmds), "0000")
                NewCmd.Parent = GetIDByName(NewGrp.Name)
                If UBound(p) > 0 Then
                    NewCmd.Actions.onclick.Type = atcURL
                    NewCmd.Actions.onclick.url = p(1)
                    NewCmd.nTextColor = 0
                    NewCmd.nBackColor = 12632256
                End If
            End If
            AddMenuCommand GetCmdParams(NewCmd)
        Else
            If nl > level + 1 Then
                MenuCmds(UBound(MenuCmds)).Actions.onclick.Type = atcNone
                MenuCmds(UBound(MenuCmds)).Actions.onmouseover.Type = atcCascade
                MenuCmds(UBound(MenuCmds)).Actions.onmouseover.TargetMenu = UBound(MenuGrps) + 1
                l = Import_ApPopUpMenu_AddGroup(lines, l - 1)
            Else
                If nl = 0 Then
                    Import_ApPopUpMenu_AddGroup = Import_ApPopUpMenu_AddGroup(lines, l) - 1
                    Exit Function
                Else
                    level = level - 1
                    Import_ApPopUpMenu_AddGroup = l - 1
                    Exit Function
                End If
            End If
        End If
    Next l
    
    Import_ApPopUpMenu_AddGroup = l - 1

End Function

Private Function Import_ApPopUpMenu_pars(line As String, level As Integer) As String()

    Import_ApPopUpMenu_pars = Split(Mid(line, 2 + level, Len(line) - (3 + level)), ",")

End Function

Private Function Import_ApPopUpMenu_GetLevel(s As String) As Integer

    Dim K As Integer
    Dim g As Integer

    For K = 2 To Len(s)
        If Mid(s, K, 1) <> "|" Then Exit For
        g = g + 1
    Next K
    
    Import_ApPopUpMenu_GetLevel = g

End Function

Private Sub Import_hierMenus4(FileName As String)

    Dim sCode As String
    Dim arrays() As String
    Dim p() As String
    Dim a As Integer
    Dim i As Integer
    Dim mPar() As String
    Dim cPar() As String
    Dim Name As String
    Dim NewCmd As MenuCmd
    Dim NewGrp As MenuGrp
    Dim c As String
    Dim g As String
    Dim HasParams As Boolean
    
    On Error Resume Next
    
    sCode = LoadFile(FileName)
    arrays = Split(sCode, "HM_Array")
    
    For a = 1 To UBound(arrays)
        Name = Left(arrays(a), InStr(arrays(a), " ") - 1)
        p = Split(arrays(a), "[")
        mPar = Split(Left(p(2), InStr(p(2), "]") - 1), ",")
        HasParams = UBound(mPar) <> -1
        
        If HasParams Then
            For i = 0 To UBound(mPar)
                If InStr(mPar(i), vbCrLf) Then
                    mPar(i) = Mid(mPar(i), InStr(mPar(i), vbCrLf) + 2)
                End If
                mPar(i) = Replace(mPar(i), """", "")
            Next i
        End If
        
        NewGrp = TemplateGroup
        NewGrp.Name = "HM_Array" & Name
        NewGrp.IncludeInToolbar = False
        If HasParams Then NewGrp.fWidth = Val(mPar(0))
        AddMenuGroup GetGrpParams(NewGrp)
        
        If InStr(Name, "_") Then
            c = Mid(Name, InStrRev(Name, "_") + 1)
            g = Left(Name, InStrRev(Name, "_") - 1)
            i = GetIDByName("HM_Array" + g + "_cmd" & c)
            
            MenuCmds(i).Actions.onmouseover.Type = atcCascade
            MenuCmds(i).Actions.onmouseover.TargetMenu = GetIDByName(NewGrp.Name)
        End If
        
        For i = 3 To UBound(p)
            NewCmd = TemplateCommand
            NewCmd.Name = NewGrp.Name + "_cmd" & (i - 2)
            
            If HasParams Then
                NewCmd.nTextColor = HTMLColor2Long(mPar(3), NewCmd.nTextColor)
                NewCmd.hTextColor = HTMLColor2Long(mPar(4), NewCmd.hTextColor)
                NewCmd.nBackColor = HTMLColor2Long(mPar(5), NewCmd.nBackColor)
                NewCmd.hBackColor = HTMLColor2Long(mPar(6), NewCmd.hBackColor)
            End If
            
            cPar = Split(Left(p(i), InStr(p(i), "]") - 1), ",")
            NewCmd.Caption = Replace(cPar(0), """", "")
            
            NewCmd.Actions.onclick.Type = atcURL
            NewCmd.Actions.onclick.url = Replace(cPar(1), """", "")
            
            AddMenuCommand GetCmdParams(NewCmd)
        Next i
    Next a

End Sub

Private Function HTMLColor2Long(c As String, alt As Long) As Long

    If Val(c) = 0 And c <> "" Then
        HTMLColor2Long = Hex2Dec(Switch(c = "white", "FFFFFF", _
                                        c = "blue", "FF0000", _
                                        c = "red", "0000FF", _
                                        c = "yellow", "00FFFF", _
                                        c = "green", "00FF00", _
                                        c = "black", "000000", _
                                        c = "gray", "808080", _
                                        c = "darkgray", "A9A9A9", _
                                        c = "lightgray", "D3D3D3"))
    Else
        If InStr(c, "#") Then
            HTMLColor2Long = Hex2Dec(Mid(c, 2))
        Else
            HTMLColor2Long = alt
        End If
    End If

End Function
