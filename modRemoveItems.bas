Attribute VB_Name = "modRemoveItems"
Option Explicit

Private tg() As String

Public Sub RemoveGroupElement(ByVal g As Integer, Recursive As Boolean, Optional Init As Boolean)

    Dim i As Integer
    Dim t As Integer
    Dim j As Integer
    
    If Init Then ReDim tg(0)
    
    With Project
        For t = 1 To UBound(.Toolbars)
            With .Toolbars(t)
                For i = 1 To UBound(.Groups)
                    If MenuGrps(g).Name = .Groups(i) Then
                        For j = i To UBound(.Groups) - 1
                            .Groups(j) = .Groups(j + 1)
                        Next j
                        ReDim Preserve .Groups(UBound(.Groups) - 1)
                        Exit For
                    End If
                Next i
            End With
        Next t
    End With
    
    UnSubGroup g
    
    For i = UBound(MenuCmds) To 1 Step -1
        If MenuCmds(i).Parent = g Then RemoveCommandElement i, Recursive
    Next i
    
    FixParentIDs g
    FixGroupsIDs g
    
    For i = g To UBound(MenuGrps) - 1
        MenuGrps(i) = MenuGrps(i + 1)
    Next i
    
    frmMain.tvMenus.SelectedItem.key = ""
    
    For i = 1 To frmMain.tvMenus.Nodes.Count
        If LenB(frmMain.tvMenus.Nodes(i).key) <> 0 Then
            If IsGroup(frmMain.tvMenus.Nodes(i).key) And GetID(frmMain.tvMenus.Nodes(i)) > g Then
                frmMain.tvMenus.Nodes(i).key = "G" & GetID(frmMain.tvMenus.Nodes(i)) - 1
            End If
        End If
    Next i
    
    ReDim Preserve MenuGrps(UBound(MenuGrps) - 1)
    
    If Init Then RemoveCascading
    
End Sub

Public Sub RemoveCommandElement(ByVal c As Integer, Recursive As Boolean, Optional Init As Boolean)

    Dim i As Integer
    Dim cc As Integer
    Dim g As Integer
    
    If Init Then ReDim tg(0)
    
    If Recursive Then
        If MenuCmds(c).Actions.OnMouseOver.Type = atcCascade Then
            g = MenuCmds(c).Actions.OnMouseOver.TargetMenu
        End If
        If MenuCmds(c).Actions.OnClick.Type = atcCascade Then
            g = MenuCmds(c).Actions.OnClick.TargetMenu
        End If
        If MenuCmds(c).Actions.OnDoubleClick.Type = atcCascade Then
            g = MenuCmds(c).Actions.OnDoubleClick.TargetMenu
        End If
        If g <> 0 Then
            ReDim Preserve tg(UBound(tg) + 1)
            tg(UBound(tg)) = MenuGrps(g).Name
        End If
    End If

    frmMain.tvMenus.Nodes(IIf(MenuCmds(c).Name = "[SEP]", "S", "C") & c).key = ""
    For i = c To UBound(MenuCmds) - 1
        MenuCmds(i) = MenuCmds(i + 1)
    Next i
    
    cc = c + 1
    Do Until cc > UBound(MenuCmds)
        For i = 1 To frmMain.tvMenus.Nodes.Count
            If LenB(frmMain.tvMenus.Nodes(i).key) <> 0 Then
                If Not IsGroup(frmMain.tvMenus.Nodes(i).key) And GetID(frmMain.tvMenus.Nodes(i)) = cc Then
                    frmMain.tvMenus.Nodes(i).key = Left$(frmMain.tvMenus.Nodes(i).key, 1) & GetID(frmMain.tvMenus.Nodes(i)) - 1
                    cc = cc + 1
                    If cc > UBound(MenuCmds) Then Exit For
                End If
            End If
        Next i
    Loop
    
    ReDim Preserve MenuCmds(UBound(MenuCmds) - 1)
    
    If Init Then RemoveCascading
    
End Sub

Private Sub RemoveCascading()

    Dim i As Integer
    Dim g As Integer
    
ReStart:
    For i = 1 To UBound(tg)
        If LenB(tg(i)) <> 0 Then
            g = GetIDByName(tg(i))
            If g > 0 Then
                tg(i) = ""
                frmMain.SelectItem frmMain.tvMenus.Nodes("G" & g)
                RemoveGroupElement g, True
                GoTo ReStart:
            End If
        End If
    Next i

End Sub

Private Sub UnSubGroup(ByVal g As Integer)

    Dim c As Integer
    
    For c = 1 To UBound(MenuCmds)
        With MenuCmds(c).Actions.OnClick
            If .Type = atcCascade And .TargetMenu = g Then .Type = atcNone
        End With
        With MenuCmds(c).Actions.OnMouseOver
            If .Type = atcCascade And .TargetMenu = g Then .Type = atcNone
        End With
        With MenuCmds(c).Actions.OnDoubleClick
            If .Type = atcCascade And .TargetMenu = g Then .Type = atcNone
        End With
    Next c

End Sub

Private Sub FixParentIDs(ByVal FromItem As Integer)

    Dim c As Integer
    
    For c = 1 To UBound(MenuCmds)
        If MenuCmds(c).Parent > FromItem Then
            MenuCmds(c).Parent = MenuCmds(c).Parent - 1
        End If
        With MenuCmds(c).Actions
            If .OnClick.Type = atcCascade And .OnClick.TargetMenu > FromItem Then
                .OnClick.TargetMenu = .OnClick.TargetMenu - 1
            End If
            If .OnMouseOver.Type = atcCascade And .OnMouseOver.TargetMenu > FromItem Then
                .OnMouseOver.TargetMenu = .OnMouseOver.TargetMenu - 1
            End If
            If .OnDoubleClick.Type = atcCascade And .OnDoubleClick.TargetMenu > FromItem Then
                .OnDoubleClick.TargetMenu = .OnDoubleClick.TargetMenu - 1
            End If
        End With
    Next c

End Sub

Private Sub FixGroupsIDs(ByVal FromItem As Integer)

    Dim g As Integer
    
    For g = 1 To UBound(MenuGrps)
        With MenuGrps(g).Actions
            If .OnClick.Type = atcCascade And .OnClick.TargetMenu > FromItem Then
                .OnClick.TargetMenu = .OnClick.TargetMenu - 1
            End If
            If .OnMouseOver.Type = atcCascade And .OnMouseOver.TargetMenu > FromItem Then
                .OnMouseOver.TargetMenu = .OnMouseOver.TargetMenu - 1
            End If
            If .OnDoubleClick.Type = atcCascade And .OnDoubleClick.TargetMenu > FromItem Then
                .OnDoubleClick.TargetMenu = .OnDoubleClick.TargetMenu - 1
            End If
        End With
    Next g

End Sub
