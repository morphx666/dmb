Attribute VB_Name = "modPresetsAPI"
Option Explicit

Public Enum PresetInfoConstants
    piTitle = 1
    piAuthor = 2
    piComments = 3
    piCategory = 4
End Enum

Public Enum PresetWorkingModeConstants
    pwmNormal = 1
    pwmApplyStyle = 2
    pwmSubmit = 3
End Enum

Public PresetWorkingMode As PresetWorkingModeConstants

Public Function GetPresetProperty(FileName As String, Property As PresetInfoConstants) As String

    Dim s() As String
    Dim sCode As String
    
    On Error Resume Next
    
    sCode = LoadPresetFile(FileName)
    s = Split(sCode, "|")
    
    Select Case Property
        Case piTitle: sCode = s(0)
        Case piAuthor: sCode = s(1)
        Case piComments: sCode = s(2)
        Case piCategory: sCode = s(3)
    End Select
    
    If InStr(sCode, "::[/]::") > 0 Then
        sCode = Left(sCode, InStr(sCode, "::[/]::") - 1)
    End If
    
    GetPresetProperty = sCode

End Function

Public Sub CleanPresetsDirs(Optional KillTmpPath As Boolean = False)

    Dim dPathTMP As String
    
    On Error Resume Next
    
    dPathTMP = TempPath + "Presets\tmp\"

    Kill dPathTMP + "*.*"
    If KillTmpPath Then RmDir dPathTMP

End Sub

#If ISCOMP = 0 Then

Public Sub CompressPreset(Title As String, Author As String, Comments As String, Category As String, flbPresets As FileListBox)

    Dim oProject As ProjectDef
    Dim oGroup() As MenuGrp
    Dim oCommand() As MenuCmd
    Dim dPath As String
    Dim dPathTMP As String
    Dim i As Integer
    Dim sStr As String
    Dim tFile As String
    
    On Error Resume Next
    
    oProject = Project
    oGroup = MenuGrps
    oCommand = MenuCmds
    
    CleanPresetsDirs

    dPath = AppPath + "Presets\"
    dPathTMP = dPath + "tmp\"
    MkDir dPath + "tmp\"
    
    For i = 1 To UBound(MenuGrps)
        With MenuGrps(i)
            With .Actions.OnClick
                If .Type = atcNewWindow Or .Type = atcURL Then .Type = atcNone
                .TargetFrame = ""
                .url = ""
                .WindowOpenParams = ""
            End With
            With .Actions.OnMouseOver
                If .Type = atcNewWindow Or .Type = atcURL Then .Type = atcNone
                .TargetFrame = ""
                .url = ""
                .WindowOpenParams = ""
            End With
            With .Actions.OnDoubleClick
                .TargetFrame = ""
                .url = ""
                .WindowOpenParams = ""
            End With
        End With
    Next i
    
    For i = 1 To UBound(MenuCmds)
        With MenuCmds(i)
            With .Actions.OnClick
                If .Type = atcNewWindow Or .Type = atcURL Then .Type = atcNone
                .TargetFrame = ""
                .url = ""
                .WindowOpenParams = ""
            End With
            With .Actions.OnMouseOver
                If .Type = atcNewWindow Or .Type = atcURL Then .Type = atcNone
                .TargetFrame = ""
                .url = ""
                .WindowOpenParams = ""
            End With
            With .Actions.OnDoubleClick
                If .Type = atcNewWindow Or .Type = atcURL Then .Type = atcNone
                .TargetFrame = ""
                .url = ""
                .WindowOpenParams = ""
            End With
        End With
    Next i
    
    With Project
        .Name = Title
        .FileName = dPathTMP + Title + ".dmb"
        With .UserConfigs(0)
            .RootWeb = dPathTMP
            .CompiledPath = .RootWeb
            .ImagesPath = .RootWeb
            .Frames.UseFrames = False
        End With
        .DefaultConfig = 0
        .CodeOptimization = cocDEBUG
        .CompilehRefFile = False
        .CompileIECode = True
        .CompileNSCode = False
        .JSFileName = "menu"
    End With
    frmMain.SaveMenu False
    frmMain.DoCompile True, , dPathTMP
    
    CopyProjectImages dPathTMP
    
    sStr = Title + "|" + Author + "|" + Comments + "|" + Category + "::[/]::"
    With flbPresets
        .Path = dPathTMP
        .Pattern = "*.*"
        .Refresh
        
        For i = 0 To .ListCount - 1
            tFile = dPathTMP + .List(i)
            If IsBin(tFile) Then
                sStr = sStr + .List(i) + "::[/]::" + LoadImageFile(tFile) + "::[/]::"
            Else
                sStr = sStr + .List(i) + "::[/]::" + LoadFile(tFile) + "::[/]::"
            End If
        Next i
    End With
    
    Kill dPath + Title + ".dpp"
    modZlib.Compress dPath + Title + ".dpp", sStr
    
    CleanPresetsDirs
    
    Project = oProject
    MenuGrps = oGroup
    MenuCmds = oCommand

End Sub

#End If

Private Function IsBin(FileName As String) As Boolean

    Dim fExt As String
    
    fExt = "*." + LCase(GetFileExtension(FileName))
    IsBin = (InStr(SupportedImageFiles, fExt) > 0) Or (InStr(SupportedCursorFiles, fExt) > 0) Or (InStr(SupportedImageFilesFlash, fExt) > 0)

End Function

Private Function LoadPresetFile(FileName As String) As String

    Dim sCode As String
    
    sCode = LoadFile(FileName)
    If InStr(sCode, "::[/]::") = 0 Then
        sCode = modZlib.UnCompress(FileName)
    End If
    
    LoadPresetFile = sCode

End Function

Public Sub UncompressPreset(pName As String)

    Dim dPath As String
    Dim dPathTMP As String
    Dim s() As String
    Dim i As Integer
    
    Dim fName As String
    Dim fCode As String
    
    On Error Resume Next
    
    CleanPresetsDirs
    
    dPath = AppPath + "Presets\"
    dPathTMP = TempPath + "Presets\tmp\"
    
    MkDir2 dPathTMP
    
    s = Split(LoadPresetFile(dPath + pName + ".dpp"), "::[/]::")
    
    For i = 1 To UBound(s) - 1 Step 2
        fName = s(i + 0)
        fCode = s(i + 1)
        
        If IsBin(fName) Then
            SaveImageFile dPathTMP + fName, fCode
        Else
            SaveFile dPathTMP + fName, fCode
        End If
    Next i

End Sub
