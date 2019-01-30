Attribute VB_Name = "modCustomSets"
Option Explicit

Public Enum csdAppliesToConstants
    atcGroups = 0
    atcCommands = 1
    atcGroupsAndCommands = 2
    atcSeparators = 3
End Enum

Public Type CustomSetsDef
    Name As String
    AppliesTo As csdAppliesToConstants
    
    ColorNormalText() As String
    ColorNormalBack() As String
    ColorOverText() As String
    ColorOverBack() As String
    
    nFontName() As String
    nFontBold() As String
    nFontItalic() As String
    nFontUnderline() As String
    nFontSize() As String
    
    oFontName() As String
    oFontBold() As String
    oFontItalic() As String
    oFontUnderline() As String
    oFontSize() As String
    
    Alignment() As String
    
    Cursor() As String
    
    ActionType() As String
    
    TargetFrame() As String
    
    ColorLine() As String
    ColorBack() As String
    
    SeparatorLength() As String
    
    ImageLeftNormal() As String
    ImageLeftOver() As String
    ImageBackNormal() As String
    ImageBackOver() As String
    ImageRightNormal() As String
    ImageRightOver() As String
End Type
Public CustomSets() As CustomSetsDef

Public SelCustomSet As Integer

Public Sub ApplyStyle(ByVal stdOP As Integer, ByVal cusOP As Integer, SelNodes() As String)

    Dim selCusSet As CustomSetsDef
    Dim i As Integer
    
    dmbClipboard.CmdContents = MenuCmds(GetID)
    dmbClipboard.ObjSrc = docCommand
    ReDim dmbClipboard.CustomSel(0)
    
    If stdOP > 0 Then
        Select Case stdOP
            Case 1
                Add2DMBCB MenuCmds(i).Name
            Case 2
                For i = 1 To UBound(MenuCmds)
                    If MenuCmds(i).Parent = MenuCmds(GetID).Parent Then
                        Add2DMBCB MenuCmds(i).Name
                    End If
                Next i
            Case 3
                For i = 1 To UBound(MenuCmds)
                    If BelongsToToolbar(i, False) = BelongsToToolbar(GetID, False) Then
                        Add2DMBCB MenuCmds(i).Name
                    End If
                Next i
            Case 4
                For i = 1 To UBound(MenuCmds)
                    Add2DMBCB MenuCmds(i).Name
                Next i
        End Select
    End If
    
    If cusOP > 0 Then
        selCusSet = CustomSets(cusOP)
        
        For i = 1 To UBound(MenuCmds)
            If HasValue(selCusSet.ColorNormalText, GetRGB(MenuCmds(i).nTextColor, True)) Then Add2DMBCB MenuCmds(i).Name
            If HasValue(selCusSet.ActionType, IIf(MenuCmds(i).Actions.onclick.Type = atcCascade, "none", "")) Then Add2DMBCB MenuCmds(i).Name
            If HasValue(selCusSet.ActionType, IIf(MenuCmds(i).Actions.onclick.Type = at, "none", "")) Then Add2DMBCB MenuCmds(i).Name
            If HasValue(selCusSet.ActionType, IIf(MenuCmds(i).Actions.onclick.Type = atcNone, "none", "")) Then Add2DMBCB MenuCmds(i).Name
        Next i
    End If
    
    SelectiveCopy False
        
    For i = 0 To UBound(SelNodes)
        With frmSelectiveCopyPaste.tvProperties
            .Nodes(SelNodes(i)).Checked = True
            frmSelectiveCopyPaste.tvProperties_NodeCheck .Nodes(SelNodes(i))
        End With
    Next i
    
    SelectivePaste 3

End Sub

Private Function HasValue(a() As String, v As String) As Boolean

    If v <> "" Then
        HasValue = InStr(("|" + Join(a, "|") + "|"), "|" + v + "|")
    End If

End Function

Private Sub Add2DMBCB(a() As String, v As String)

    With dmbClipboard
        ReDim Preserve .CustomSel(UBound(.CustomSel) + 1)
        .CustomSel(UBound(.CustomSel)) = v
    End With

End Sub
