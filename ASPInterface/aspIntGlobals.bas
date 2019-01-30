Attribute VB_Name = "aspIntGlobals"
Option Explicit

Private Sub Main()

    Dim i As Integer
    
    On Error GoTo dspErr
    
    cSep = Chr(255) + Chr(255)
    
    SetTemplateDefaults
    
    Set FloodPanel.PictureControl = frmMain.picFlood
    Dim UIObjects(1 To 2) As Object
    Set UIObjects(1) = frmMain
    Set UIObjects(2) = FloodPanel
    SetUI UIObjects
    
    Dim VarObjects(1 To 6) As Variant
    VarObjects(1) = GetSetting("DMB", "RegInfo", "InstallPath")
    VarObjects(2) = ""
    VarObjects(3) = ""
    VarObjects(4) = GetTEMPPath
    VarObjects(5) = cSep
    VarObjects(6) = nwdPar
    SetVars VarObjects
    
    Erase MenuGrps: ReDim MenuGrps(0)
    Erase MenuCmds: ReDim MenuCmds(0)
    
    With Project
        ReDim .UserConfigs(0)
        With .UserConfigs(0)
            .Name = "Local"
            .Description = "Publish the menus on the local machine without the use of a web server"
            .Type = ctcLocal
        End With
    End With
    
    Exit Sub
    
dspErr:
    ReError

End Sub

Public Function LoadMenu(File As String) As Boolean

    Dim sStr As String
    Dim Ans As Integer
    Dim nLines As Integer
    Dim cLine As Integer
        
    On Error GoTo chkError
    
    Project = GetProjectProperties(File)
    
    Erase MenuGrps: ReDim MenuGrps(0)
    Erase MenuCmds: ReDim MenuCmds(0)
    
    If EOF(ff) Then GoTo ExitSub
    Line Input #ff, sStr
    Do Until EOF(ff) Or sStr = "[RSC]"
        AddMenuGroup Mid$(sStr, 4)
        Do Until EOF(ff) Or sStr = "[RSC]"
            Line Input #ff, sStr
            If Left$(sStr, 3) = "[C]" Then
                AddMenuCommand Mid$(sStr, 6), True
            Else
                Exit Do
            End If
        Loop
        If EOF(ff) Then Exit Do
    Loop
    
    LoadMenu = True
ExitSub:
    Close ff
    
    FloodPanel.Value = 0
   
    Exit Function
    
chkError:
    ReError

End Function

Private Sub GetPrgPrefs()
    
    USER = GetSetting(App.EXEName, "RegInfo", "User", "DEMO")
    COMPANY = GetSetting(App.EXEName, "RegInfo", "Company", "DEMO")
    USERSN = GetSetting(App.EXEName, "RegInfo", "SerialNumber", "")
    
    With Preferences
        .AutoRecover = GetSetting(App.EXEName, "Preferences", "AutoRecover", True)
        .OpenLastProject = GetSetting(App.EXEName, "Preferences", "OpenLastProject", True)
        .SepHeight = GetSetting(App.EXEName, "Preferences", "SepHeight", 13)
        .ShowNag = GetSetting(App.EXEName, "Preferences", "ShowNag", True)
        .ShowWarningAddInEditor = GetSetting(App.EXEName, "Preferences", "ShowWarningAIE", True)
        .CommandsInheritance = GetSetting(App.EXEName, "Preferences", "CmdInh", icFirst)
        .GroupsInheritance = GetSetting(App.EXEName, "Preferences", "GrpInh", icFirst)
        .UseLivePreview = GetSetting(App.EXEName, "Preferences", "UseLivePreview", True)
        .DisableUndoRedo = GetSetting(App.EXEName, "Preferences", "DisableUR", False)
        .ImgSpace = Val(GetSetting(App.EXEName, "Preferences", "ImgSpace", 4))
    End With
    
End Sub

Public Sub ReError()

    Dim objContext As ObjectContext
    Dim errMsg As String
    
    errMsg = "Error " & Err.Number & ": " & Err.Description
    
    Set objContext = GetObjectContext()
    If objContext Is Nothing Then
        MsgBox errMsg, vbOKOnly + vbCritical, "DMB ASP Interface Error"
    Else
        objContext("Response").Write ("DMB ASP Interface Error" & vbCrLf & errMsg)
        Set objContext = Nothing
    End If
    
End Sub
