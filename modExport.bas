Attribute VB_Name = "modExport"
Option Explicit

Dim gApplyStyles As exHTMLStylesConstants
Dim gGrpClass As String
Dim gCmdClass As String
Dim gIdent As Integer
Dim gImgWidth As Integer
Dim gImgHeight As Integer
Dim gImgC As String
Dim gImgE As String
Dim gImgN As String
Dim gImgM As String
Dim gFileName As String
Dim gExpLinks As Boolean
Dim gSingleSel As Boolean
Dim cIdent As Integer
Dim t As Integer
Dim SubLevel As Integer

Const NullURL = "javascript:void(0);"

#If STANDALONE = 0 Then

Public Sub ExportAsHTML(Title As String, _
                        Description As String, _
                        FileName As String, _
                        ApplyStyles As exHTMLStylesConstants, _
                        CSSFile As String, _
                        gCls As String, _
                        cCls As String, _
                        TreeLike As Boolean, _
                        cPic As String, _
                        ePic As String, _
                        nPic As String, _
                        Identation As Integer, _
                        imgPath As String, _
                        imgW As Integer, _
                        imgH As Integer, _
                        ExpItemsLinks As Boolean, _
                        SingleSel As Boolean, _
                        IncludeExpCol As Boolean, _
                        ExpAllStr As String, _
                        ColAllStr As String, _
                        ByVal ExpColPlacement As exHTMLExpColPlacementConstants, _
                        XHTMLCompliant As Boolean)

    Dim html As String
    Dim i As Integer
    Dim css As String
    Dim AnchorStr As String
       
    On Error Resume Next
    
    SubLevel = 0
    
    For i = 1 To UBound(MenuGrps)
        MenuGrps(i).Name = "grp" + MenuGrps(i).Name
        MenuGrps(i).caption = Replace(MenuGrps(i).caption, " & ", " &amp; ")
    Next i
    
    For i = 1 To UBound(MenuCmds)
        MenuCmds(i).caption = Replace(MenuCmds(i).caption, " & ", " &amp; ")
    Next i
    
    CompileProject MenuGrps, MenuCmds, Project, Preferences, params, False, True
    
    If XHTMLCompliant Then
        html = "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">" + vbCrLf
        html = html + "<html xmlns=""http://www.w3.org/1999/xhtml"">" + vbCrLf
    Else
        html = "<!DOCTYPE html PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"" ""http://www.w3.org/TR/html4/loose.dtd"">" + vbCrLf
        html = html + "<html>" + vbCrLf
    End If
    
    html = html + "<head>" + vbCrLf
    html = html + "<meta http-equiv=""Content-Type"" content=""text/html; charset=windows-1252"" />" + vbCrLf
    html = html + "<meta name=""description"" content=""" + RemoveHTMLCode(Description) + """ />" + vbCrLf
    html = html + "<meta name=""generator"" content=""DHTML Menu Builder " + DMBVersion + """ />" + vbCrLf
    html = html + "<title>" + RemoveHTMLCode(Title) + "</title>" + vbCrLf
    html = html + "<style type=""text/css"">" + vbCrLf
    
    gIdent = Identation
    gImgWidth = imgW
    gImgHeight = imgH
    gImgC = SetSlashDir(GetSmartRelPath(FileName, imgPath) + GetFileName(cPic), sdFwd)
    gImgE = SetSlashDir(GetSmartRelPath(FileName, imgPath) + GetFileName(ePic), sdFwd)
    gImgN = SetSlashDir(GetSmartRelPath(FileName, imgPath) + GetFileName(nPic), sdFwd)
    gImgM = SetSlashDir(GetSmartRelPath(FileName, imgPath) + "m.gif", sdFwd)
    gFileName = FileName
    gExpLinks = ExpItemsLinks
    gSingleSel = SingleSel
    gApplyStyles = ApplyStyles
    Select Case gApplyStyles
        Case ascNone
            html = html + "a {text-decoration: none;}" + vbCrLf
            html = html + "a:Hover {background-color: #C0C0C0;}" + vbCrLf
        Case ascProject
            t = UBound(MenuGrps) + UBound(MenuCmds)
            FloodPanel.caption = "Generating CSS Code..."
            For i = 1 To UBound(MenuGrps)
                FloodPanel.Value = i / t * 100
                With MenuGrps(i)
                    If .Compile Then
                        If .Actions.onclick.Type = atcURL Or .Actions.onclick.Type = atcNewWindow Then
                            AnchorStr = "A"
                        Else
                            AnchorStr = ""
                        End If
                        css = css + AnchorStr + ".grp" + .Name + " {font-family: " + .DefNormalFont.FontName + "; font-size: " & .DefNormalFont.FontSize & "; color: " + GetRGB(.nTextColor) + "; background-color: " + GetRGB(.nBackColor, True) + "; font-weight: " + IIf(.DefNormalFont.FontBold, "bold", "normal") + "; font-style: " + IIf(.DefNormalFont.FontItalic, "italic", "normal") + "; text-decoration: " + IIf(.DefNormalFont.FontUnderline, "underline", "none") + "}" + vbCrLf
                        If .disabled Then
                            css = css + AnchorStr + ".grp" + .Name + ":Hover {font-family: " + .DefNormalFont.FontName + "; font-size: " & .DefNormalFont.FontSize & "; color: " + GetRGB(.nTextColor) + "; background-color: " + GetRGB(.nBackColor, True) + "; font-weight: " + IIf(.DefNormalFont.FontBold, "bold", "normal") + "; font-style: " + IIf(.DefNormalFont.FontItalic, "italic", "normal") + "; text-decoration: " + IIf(.DefNormalFont.FontUnderline, "underline", "none") + "}" + vbCrLf
                        Else
                            css = css + AnchorStr + ".grp" + .Name + ":Hover {font-family: " + .DefNormalFont.FontName + "; font-size: " & .DefNormalFont.FontSize & "; color: " + GetRGB(.hTextColor) + "; background-color: " + GetRGB(.hBackColor, True) + "; font-weight: " + IIf(.DefHoverFont.FontBold, "bold", "normal") + "; font-style: " + IIf(.DefHoverFont.FontItalic, "italic", "normal") + "; text-decoration: " + IIf(.DefHoverFont.FontUnderline, "underline", "none") + "}" + vbCrLf
                        End If
                    End If
                End With
            Next i
            FloodPanel.caption = "Generating CSS Code..."
            For i = 1 To UBound(MenuCmds)
                FloodPanel.Value = i / t * 100
                With MenuCmds(i)
                    If .Compile Then
                        If .Actions.onclick.Type = atcURL Or .Actions.onclick.Type = atcNewWindow Then
                            AnchorStr = "A"
                        Else
                            AnchorStr = ""
                        End If
                        css = css + AnchorStr + ".cmd" + .Name + " {font-family: " + .NormalFont.FontName + "; font-size: " & .NormalFont.FontSize & "; color: " + GetRGB(.nTextColor) + "; background-color: " + GetRGB(.nBackColor, True) + "; font-weight: " + IIf(.NormalFont.FontBold, "bold", "normal") + "; font-style: " + IIf(.NormalFont.FontItalic, "italic", "normal") + "; text-decoration: " + IIf(.NormalFont.FontUnderline, "underline", "none") + "}" + vbCrLf
                        If .disabled Then
                            css = css + AnchorStr + ".cmd" + .Name + ":Hover {font-family: " + .NormalFont.FontName + "; font-size: " & .NormalFont.FontSize & "; color: " + GetRGB(.nTextColor) + "; background-color: " + GetRGB(.nBackColor, True) + "; font-weight: " + IIf(.NormalFont.FontBold, "bold", "normal") + "; font-style: " + IIf(.NormalFont.FontItalic, "italic", "normal") + "; text-decoration: " + IIf(.NormalFont.FontUnderline, "underline", "none") + "}" + vbCrLf
                        Else
                            css = css + AnchorStr + ".cmd" + .Name + ":Hover {font-family: " + .NormalFont.FontName + "; font-size: " & .NormalFont.FontSize & "; color: " + GetRGB(.hTextColor) + "; background-color: " + GetRGB(.hBackColor, True) + "; font-weight: " + IIf(.HoverFont.FontBold, "bold", "normal") + "; font-style: " + IIf(.HoverFont.FontItalic, "italic", "normal") + "; text-decoration: " + IIf(.HoverFont.FontUnderline, "underline", "none") + "}" + vbCrLf
                        End If
                    End If
                End With
            Next i
            html = html + OptimizeCSS(css)
        Case ascCSS
            gGrpClass = Mid(gCls, 2)
            gCmdClass = Mid(cCls, 2)
    End Select
    html = html + "li {list-style-type: disc;}" + vbCrLf
    
    html = html + "</style>" + vbCrLf
    If gApplyStyles = ascCSS Then html = html + "<link rel=""stylesheet"" type=""text/css"" href=""" + CSSFile + """>"
    html = html + "</head>" + vbCrLf
    html = html + "<body>" + vbCrLf
    
    html = html + "<div style=""font-weight:bold; font-size:x-large; font-family:" + MenuGrps(1).DefNormalFont.FontName + """>" + vbCrLf
    html = html + Title + "<br>" + vbCrLf
    html = html + Description + vbCrLf
    html = html + "</div>"
    html = html + "<p>&nbsp;</p>" + vbCrLf
    
    If TreeLike Then
        If IncludeExpCol And (ExpColPlacement = ecpcTop Or ExpColPlacement = ecpcBoth) Then
            html = html + "<a href=""javascript:toggleAll(1)"">" + ExpAllStr + "</a> | <a href=""javascript:toggleAll(0)"">" + ColAllStr + "</a><br>"
        End If
        html = html + TreeHTML
        If Not IncludeExpCol Then RemoveFunction html, "toggleAll"
        If IncludeExpCol And (ExpColPlacement = ecpcBottom Or ExpColPlacement = ecpcBoth) Then
            html = html + "<br><br><a href=""javascript:toggleAll(1)"">" + ExpAllStr + "</a> | <a href=""javascript:toggleAll(0)"">" + ColAllStr + "</a>"
        End If
        FileCopy cPic, imgPath + GetFileName(cPic)
        FileCopy ePic, imgPath + GetFileName(ePic)
        FileCopy nPic, imgPath + GetFileName(nPic)
        FileCopy AppPath + "exhtml\m.gif", imgPath + "m.gif"
    Else
        html = html + StdHTML
    End If
    
    html = html + "</body>" + vbCrLf
    html = html + "</html>" + vbCrLf
    
    FloodPanel.Value = 0
    
    If XHTMLCompliant Then html = ApplyXHTMLRules(html)
    
    SaveFile FileName, html
    
    For i = 1 To UBound(MenuGrps)
        MenuGrps(i).Name = Mid(MenuGrps(i).Name, 4)
    Next i

End Sub

Private Function ApplyXHTMLRules(html As String) As String
   
    html = CloseTag(html, "meta")
    html = CloseTag(html, "img")
    html = CloseTag(html, "br")
    
    Dim p1 As Long
    p1 = 1
    Do
        p1 = InStr(p1, html, "javascript"">")
        If p1 = 0 Then Exit Do
        html = Mid(html, 1, p1 + 11) + vbCrLf + "/* <![CDATA[ */" + vbCrLf + Mid(html, p1 + 12)
        
        p1 = InStr(p1, html, "</script>")
        html = Mid(html, 1, p1 - 1) + vbCrLf + "/* ]]> */" + vbCrLf + Mid(html, p1)
        
        p1 = p1 + 1
    Loop
    
    ApplyXHTMLRules = html
    
End Function

Private Function CloseTag(html As String, tag As String) As String

    Dim p1 As Long
    Dim p2 As Long
    
    p1 = 1
    Do
        p1 = InStr(p1, html, "<" + tag)
        If p1 = 0 Then Exit Do
        p2 = InStr(p1, html, ">")
        
        html = Mid(html, 1, p2 - 1) + "/" + Mid(html, p2)
        p1 = p2
    Loop
    
    CloseTag = html
    
End Function

Private Function TreeHTML() As String

          Dim g As Integer
          Dim k As Integer
          Dim html As String
          Dim IncludeNoTBItems As Boolean
          Dim DoRender As Boolean
          Dim t As Integer
          Dim fName As String
          
   On Error GoTo TreeHTML_Error

10        cIdent = gIdent
          
20        IncludeNoTBItems = Not CreateToolbar
          
30        If CreateToolbar Then
40            For t = 1 To UBound(Project.Toolbars)
50                If Project.Toolbars(t).Compile And Project.Toolbars(t).IsTemplate = False Then
60                    For k = 1 To UBound(Project.Toolbars(t).Groups)
70                        g = GetIDByName("grp" + Project.Toolbars(t).Groups(k))
80                        With MenuGrps(g)
90                            DoRender = (Not IsSubMenu(g)) And .Compile
100                           If LenB(.caption) = 0 Then DoRender = DoRender And (GroupExpands(g) Or (.Actions.onclick.Type = atcURL))
110                           If DoRender Then
120                               fName = gfName(.Actions.onclick.TargetFrame)
130                               html = html + "<br><span>" + _
                                                  GetImgCode(IIf(GroupExpands(g), gImgC, gImgN), .Name, Not .disabled And GroupExpands(g)) + _
                                                  "<a href=""" + IIf(gExpLinks, gtURL(.Actions, .disabled), "javascript:toggleState('" + .Name + "');") + """ " + GetClass(True, .Name) + " onmousemove=""status='" + EncodeStr(PrepareCaption(.WinStatus, False)) + "';""" + IIf(GetRealLocal.Frames.UseFrames And LenB(fName) <> 0, " target=""" + gfName(.Actions.onclick.TargetFrame) + """", "") + ">" + IIf(LenB(.caption) = 0, .Name, .caption) + _
                                                  "</a></span>" + vbCrLf
                                  'html = html + "<span style=""display:none"" id=" + .Name + "BR><br></span>" + vbCrLf
140                               html = html + "<span id=""" + .Name + """ style=""display:none;"">" + vbCrLf
150                               html = html + GenCmdTreeCode(GetRealSubGroup(g))
160                               html = html + "</span>" + vbCrLf
170                           End If
180                       End With
190                   Next k
200               End If
210           Next t
220       Else
230           For g = 1 To UBound(MenuGrps)
240               With MenuGrps(g)
250                   DoRender = (Not IsSubMenu(g)) And .Compile
260                   If LenB(.caption) = 0 Then DoRender = DoRender And (GroupExpands(g) Or (.Actions.onclick.Type = atcURL))
270                   If DoRender Then
280                       fName = gfName(.Actions.onclick.TargetFrame)
290                       html = html + "<br><span>" + _
                                          GetImgCode(IIf(GroupExpands(g), gImgC, gImgN), .Name, Not .disabled And GroupExpands(g)) + _
                                          "<a href=""" + IIf(gExpLinks, gtURL(.Actions, .disabled), "javascript:toggleState('" + .Name + "');") + """ " + GetClass(True, .Name) + " onmousemove=""status='" + EncodeStr(PrepareCaption(.WinStatus, False)) + "';""" + IIf(GetRealLocal.Frames.UseFrames And LenB(fName) <> 0, " target=""" + gfName(.Actions.onclick.TargetFrame) + """", "") + ">" + IIf(LenB(.caption) = 0, .Name, .caption) + _
                                          "</a></span>" + vbCrLf
                          'html = html + "<span style=""display:none"" id=" + .Name + "BR><br></span>" + vbCrLf
300                       html = html + "<span id=""" + .Name + """ style=""display:none;"">" + vbCrLf
310                       html = html + GenCmdTreeCode(GetRealSubGroup(g))
320                       html = html + "</span>" + vbCrLf
330                   End If
340               End With
350           Next g
360       End If
          
370       html = html + "<script language=""javascript"" type=""text/javascript"">"
380       html = html + "var cImg = new Image();cImg.src = '" + gImgC + "';"
390       html = html + "var eImg = new Image();eImg.src = '" + gImgE + "';"
400       html = html + Replace(LoadFile(AppPath + "rsc/treecode.dat"), "%%SINGLESEL%%", "var ssel = " + IIf(gSingleSel, "true", "false") + ";")
410       html = html + "</script>"
          
420       TreeHTML = html

   On Error GoTo 0
   Exit Function

TreeHTML_Error:

    MsgBox "Error " & Err.number & " (" & Err.Description & ") in procedure TreeHTML of Module modExport at line " & Erl

End Function

Private Function GetImgCode(FileName As String, ItemName As String, AddToggleCode As Boolean) As String

    Dim html As String

    If LenB(FileName) <> 0 Then
        If AddToggleCode Then html = html + "<a href=""" + NullURL + """ onclick=""toggleState('" + ItemName + "')"">"
        html = html + "<img src=""" + FileName + """ width=""" & gImgWidth & """ height=""" & gImgHeight & """ align=""middle"" name=""" + ItemName + "Img"" border=""0"" alt="""">"
        If AddToggleCode Then html = html + "</a>"
    End If
    
    GetImgCode = html

End Function

Private Function OptimizeCSS(ByVal css As String) As String

    Dim l() As String
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
    Dim refCSS As String
    
    l = Split(css, vbCrLf)
ReStart:
    For i = 0 To UBound(l)
        refCSS = Mid(l(i), InStr(l(i), " ") + 1)
        For j = i - 1 To 0 Step -1
            If refCSS = Mid(l(j), InStr(l(j), " ") + 1) Then
                l(j) = Left(l(j), InStr(l(j), " ") - 1) + "," + Left(l(i), InStr(l(i), " ") - 1) + " " + refCSS
                For k = i To UBound(l) - 1
                    l(k) = l(k + 1)
                Next k
                ReDim Preserve l(UBound(l) - 1)
                GoTo ReStart
            End If
        Next j
    Next i
    
    OptimizeCSS = Join(l, vbCrLf)

End Function

Private Function GenCmdTreeCode(g As Integer)

    Dim html As String
    Dim c As Integer
    Dim IsCascade As Boolean
    Dim TargetMenu As String
    Dim fixName As String
    Dim fName As String

    For c = 1 To UBound(MenuCmds)
        With MenuCmds(c)
            If .parent = g And .Compile Then
                If .Name <> "[SEP]" And LenB(.caption) <> 0 Then
                    IsCascade = (.Actions.onmouseover.Type = atcCascade Or .Actions.onclick.Type = atcCascade Or .Actions.OnDoubleClick.Type = atcCascade)
                    
                    If IsCascade Then
                        If .Actions.onmouseover.Type = atcCascade Then
                            TargetMenu = MenuGrps(.Actions.onmouseover.TargetMenu).Name
                        Else
                            If .Actions.onclick.Type = atcCascade Then
                                TargetMenu = MenuGrps(.Actions.onclick.TargetMenu).Name
                            Else
                                If .Actions.OnDoubleClick.Type = atcCascade Then
                                    TargetMenu = MenuGrps(.Actions.OnDoubleClick.TargetMenu).Name
                                End If
                            End If
                        End If
                    
                        SubLevel = SubLevel + 1
                        fixName = .Name & SubLevel
                        html = html + "<span><br><img src=""" + gImgM + """ width=""" & cIdent & """ height=""1"" alt="""">" + _
                                        GetImgCode(gImgC, fixName, Not .disabled) + _
                                        "<a href=""" + IIf(gExpLinks, gtURL(.Actions, .disabled), "javascript:toggleState('" + fixName + "');") + """ " + GetClass(False, .Name) + " onmousemove=""status='" + EncodeStr(PrepareCaption(.WinStatus, False)) + "';"">" + .caption + _
                                        "</a></span>" + vbCrLf
                        cIdent = gIdent * (SubLevel + 1)
                        'html = html + "<span style=""display:none"" id=" + fixName + "BR><br></span>" + vbCrLf
                        html = html + "<span id=""" + fixName + """ style=""display:none"">" + vbCrLf
                        html = html + GenCmdTreeCode(GetRealSubGroup(GetIDByName(TargetMenu)))
                    Else
                        fName = gfName(.Actions.onclick.TargetFrame)
                        html = html + "<br><img src=""" + gImgM + """ width=""" & cIdent & """ height=""1"" alt="""">" + _
                                        GetImgCode(gImgN, .Name, False) + _
                                        "<a href=""" + gtURL(.Actions, .disabled) + """ " + GetClass(False, .Name) + " onmousemove=""status='" + EncodeStr(PrepareCaption(.WinStatus, False)) + "';""" + IIf(GetRealLocal.Frames.UseFrames And LenB(fName) <> 0, " target=""" + fName + """", "") + ">" + .caption + _
                                        "</a>" + vbCrLf
                    End If
                
                    If IsCascade Then
                        SubLevel = SubLevel - 1
                        cIdent = gIdent * (SubLevel + 1)
                        html = html + "</span>" + vbCrLf
                    End If
                End If
            End If
        End With
    Next c
    
    'If Left(html, 14) = "<span><br><img" Then html = "<span>" + Mid(html, 11)
    'If Left(html, 4) = "<br>" Then html = Mid(html, 5)
    GenCmdTreeCode = html

End Function

Private Function GetClass(IsGroup As Boolean, ItemName As String) As String

    Dim html As String
    
    Select Case gApplyStyles
        Case ascNone
        Case ascProject
            html = "class=""" + IIf(IsGroup, "grp", "cmd") + ItemName + """"
        Case ascCSS
            html = "class=""" + IIf(IsGroup, gGrpClass, gCmdClass) + """"
    End Select

    GetClass = html
    
End Function


#End If

Private Function gtURL(act As ActionEvents, dis As Boolean) As String

    If dis Then
        gtURL = NullURL
    Else
        gtURL = IIf((act.onclick.Type = atcURL Or act.onclick.Type = atcNewWindow) And LenB(act.onclick.url) <> 0, FixURL(act.onclick.url), NullURL)
    End If
    
    If InStr(gtURL, "%%REP%%") Then
        gtURL = SetSlashDir(Replace(gtURL, "%%REP%%", GetSmartRelPath(gFileName, GetRealLocal.RootWeb)), sdFwd)
    End If

End Function

Private Function GroupExpands(g As Integer) As Boolean

    With MenuGrps(g).Actions
        GroupExpands = (.onmouseover.Type = atcCascade Or .onclick.Type = atcCascade Or .OnDoubleClick.Type = atcCascade) And Not MenuGrps(g).disabled
    End With

End Function

Public Function StdHTML(Optional IsNoScript As Boolean = False) As String

    Dim g As Integer
    Dim k As Integer
    Dim html As String
    Dim tidx As Integer
    Dim n As Integer
    Dim prefix As String
    
    If IsNoScript Then
        prefix = ""
        html = "<noscript>"
    Else
        prefix = "grp"
        html = ""
    End If
    
    FloodPanel.caption = "Generating Standard HTML Code..."

    html = html + "<ul>" + vbCrLf
    If CreateToolbar Then
        For tidx = 1 To UBound(Project.Toolbars)
            If Project.Toolbars(tidx).IsTemplate = False Then
                t = t + UBound(Project.Toolbars(tidx).Groups)
            End If
        Next tidx
        For tidx = 1 To UBound(Project.Toolbars)
            If Project.Toolbars(tidx).Compile And Project.Toolbars(tidx).IsTemplate = False Then
                For k = 1 To UBound(Project.Toolbars(tidx).Groups)
                    n = n + 1
                    g = GetIDByName(prefix + Project.Toolbars(tidx).Groups(k))
                    FloodPanel.Value = n / t * 100
                    With MenuGrps(g)
                        If .Compile Then
                            If Not IsSubMenu(g) And (GroupExpands(g) Or (.Actions.onclick.Type = atcURL)) And Not .disabled Then
                                html = html + "<li>" + GenGrpLink(g) + IIf(LenB(.caption) <> 0, .caption, "[" + .Name + "]") + GenGrpLink(g, True) + vbCrLf
                                html = html + expCmd(GetRealSubGroup(g))
                                html = html + "</li>" + vbCrLf
                            End If
                        End If
                    End With
                Next k
            End If
        Next tidx
    Else
        t = UBound(MenuGrps)
        For g = 1 To UBound(MenuGrps)
            FloodPanel.Value = g / t * 100
            With MenuGrps(g)
                If .Compile Then
                    If Not IsSubMenu(g) And (GroupExpands(g) Or (.Actions.onclick.Type = atcURL)) And Not .disabled Then
                        html = html + "<li>" + GenGrpLink(g) + IIf(LenB(.caption) <> 0, .caption, "[" + .Name + "]") + GenGrpLink(g, True) + vbCrLf
                        html = html + expCmd(GetRealSubGroup(g))
                        html = html + "</li>" + vbCrLf
                    End If
                End If
            End With
        Next g
    End If
    html = html + "</ul>"
    
    If IsNoScript Then html = html + "</noscript>"
    html = html + vbCrLf
    
    StdHTML = html

End Function

Private Function GenGrpLink(g As Integer, Optional IsClosingTag As Boolean) As String

    Dim html As String

    With MenuGrps(g)
        If .Actions.onclick.Type = atcURL Or .Actions.onclick.Type = atcNewWindow Then
            html = SetSlashDir(Replace(FixURL(.Actions.onclick.url), "%%REP%%", GetSmartRelPath(gFileName, GetRealLocal.RootWeb)), sdFwd)
        End If
        'If html = "" Then html = NullURL
        If IsClosingTag Then
            If LenB(html) = 0 Then
                html = "</span>"
            Else
                html = "</a>"
            End If
        Else
            If LenB(html) = 0 Then
                Select Case gApplyStyles
                    Case ascNone
                        html = "<span"
                    Case ascProject
                        html = "<span class=""grp" + .Name + """"
                    Case ascCSS
                        html = "<span class="" + gGrpClass+ """""
                End Select
            Else
                Select Case gApplyStyles
                    Case ascNone
                        html = "<a href=""" + html + """"
                    Case ascProject
                        html = "<a class=""grp" + .Name + """ href=""" + html + """"
                    Case ascCSS
                        html = "<a class=""" + gGrpClass + """ href=""" + html + """"
                End Select
            End If
            If GetRealLocal.Frames.UseFrames Then
                html = html + " target=""" + gfName(.Actions.onclick.TargetFrame) + """>"
            Else
                html = html + ">"
            End If
        End If
    End With
    
    GenGrpLink = html

End Function

Private Function GenCmdLink(c As Integer, Optional IsClosingTag As Boolean) As String

    Dim html As String

    With MenuCmds(c)
        If .Actions.onclick.Type = atcURL Or .Actions.onclick.Type = atcNewWindow Then
            html = SetSlashDir(Replace(FixURL(.Actions.onclick.url), "%%REP%%", GetSmartRelPath(gFileName, GetRealLocal.RootWeb)), sdFwd)
        End If
        If LenB(html) = 0 Then html = NullURL
        If IsClosingTag Then
            html = "</a>"
        Else
            Select Case gApplyStyles
                Case ascNone
                    html = "<a href=""" + html + """"
                Case ascProject
                    html = "<a class=""cmd" + .Name + """ href=""" + html + """"
                Case ascCSS
                    html = "<a class=""" + gCmdClass + """ href=""" + html + """"
            End Select
            If GetRealLocal.Frames.UseFrames Then
                html = html + " target=""" + gfName(.Actions.onclick.TargetFrame) + """>"
            Else
                html = html + ">"
            End If
        End If
    End With
    
    GenCmdLink = html

End Function

Private Function gfName(ByVal f As String) As String

    If InStr(f, ".") Then
        gfName = Split(f, ".")(UBound(Split(f, ".")))
    Else
        gfName = f
    End If

End Function

Private Function expCmd(g As Integer) As String

    Dim html As String
    Dim c As Integer
    
    html = html + "<ul>" + vbCrLf
    For c = 1 To UBound(MenuCmds)
        With MenuCmds(c)
            If .parent = g And Not .disabled And .Compile Then
                If .Name <> "[SEP]" Then
                    html = html + "<li>" + GenCmdLink(c) + IIf(LenB(.caption) <> 0, .caption, "[" + .Name + "]") + GenCmdLink(c, True)
                    If .Actions.onmouseover.Type = atcCascade Then html = html + expCmd(.Actions.onmouseover.TargetMenu)
                    If .Actions.onclick.Type = atcCascade Then html = html + expCmd(.Actions.onclick.TargetMenu)
                    If .Actions.OnDoubleClick.Type = atcCascade Then html = html + expCmd(.Actions.OnDoubleClick.TargetMenu)
                    html = html + "</li>" + vbCrLf
                End If
            End If
        End With
    Next c
    'html = html + "<hr noshade color=#C0C0C0 align=left width=80% size=1>" + vbCrLf
    html = html + "</ul>" + vbCrLf
    
    If html = "<ul>" + vbCrLf + "</ul>" + vbCrLf Then html = ""
    
    expCmd = html

End Function

Public Sub ExportAsSitemap(FileName As String)

    Dim xml As String
    Dim i As Integer
    Dim url As String
    Dim webAddress As String

    On Error GoTo ExportAsSitemap_Error

    gFileName = GetRealLocal.RootWeb
    
    If Project.UserConfigs(Project.DefaultConfig).Type = ctcRemote Then
        webAddress = Project.UserConfigs(Project.DefaultConfig).RootWeb
    Else
        webAddress = InputBox("Please specify the web site's URL where this sitemap will be hosted:" + vbCrLf + "e.g. http://www.mywebsite.com", "Web Site Address", "")
        If webAddress = "" Then Exit Sub
    End If
    If Right(webAddress, 1) <> "/" Then webAddress = webAddress + "/"
    
    xml = "<?xml version=""1.0"" encoding=""UTF-8""?>" + vbCrLf
    
    xml = xml + "<urlset xmlns=""http://www.sitemaps.org/schemas/sitemap/0.9"">" + vbCrLf
    
    For i = 1 To UBound(MenuGrps)
        If MenuGrps(i).Actions.onclick.Type = atcURL Then
            xml = xml + "<url>" + vbCrLf
            
            url = webAddress + gtURL(MenuGrps(i).Actions, False)
            url = Replace(url, "&amp;", "#amp;")
            url = Replace(url, "&", "&amp;")
            url = Replace(url, "#amp;", "&amp;")
            
            xml = xml + "<loc>" + url + "</loc>" + vbCrLf
            xml = xml + "</url>" + vbCrLf
        End If
    Next i
    
    For i = 1 To UBound(MenuCmds)
        If MenuCmds(i).Actions.onclick.Type = atcURL Then
            xml = xml + "<url>" + vbCrLf
            
            url = webAddress + gtURL(MenuCmds(i).Actions, False)
            url = Replace(url, "&amp;", "#amp;")
            url = Replace(url, "&", "&amp;")
            url = Replace(url, "#amp;", "&amp;")
            
            xml = xml + "<loc>" + url + "</loc>" + vbCrLf
            xml = xml + "</url>" + vbCrLf
        End If
250       Next i
          
260       xml = xml + "</urlset>"

270       SaveFile FileName, xml

    MsgBox "The sitemap.xml file has been successfully generated and saved in the '" + GetFilePath(FileName) + "' folder", vbInformation + vbOKOnly, "Sitemap Generation"

   On Error GoTo 0
   Exit Sub

ExportAsSitemap_Error:

    MsgBox "Error " & Err.number & " (" & Err.Description & ") in procedure ExportAsSitemap of Module modExport"
          
End Sub
