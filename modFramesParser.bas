Attribute VB_Name = "modFramesParser"
Option Explicit

Private Frames() As CFrame
Private MainFile As String

Public Function ParseFrameset(FileName As String) As CFrame()

    On Error Resume Next
    
    Dim TopFrame As CFrame
    Dim oCaption As String
    Dim oCursor As Integer
    
    oCursor = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    #If ISCOMP = 0 Then
    If NagScreenIsVisible Then
        oCaption = frmNag.lblInfo.Caption
        frmNag.lblInfo.Caption = "Parsing Frames..."
    End If
    #End If
    
    ReDim Frames(0)
    MainFile = FileName

    TopFrame.Name = "top"
    
    StartParsing FileName, TopFrame
    
    ParseFrameset = Frames
    
    Erase Frames
    
    #If ISCOMP = 0 Then
    If NagScreenIsVisible Then
        frmNag.lblInfo.Caption = oCaption
    End If
    #End If
    Screen.MousePointer = oCursor

End Function

Private Sub StartParsing(ByVal FileName As String, pFrame As CFrame)
    
    Dim c() As String
    Dim i As Integer
    
    c = GetFrameSet(FileName)
    For i = 1 To Val(c(0))
        fsparser FileName, c(i), pFrame
    Next
    
    If GetIFrames(FileName, pFrame) Then AddNewFrame pFrame.Name, ""

End Sub

Private Function GetIFrames(ByVal FileName As String, pFrame As CFrame) As Boolean

    Dim cCode As String
    Dim p1 As Long
    Dim p2 As Long
    Dim Exists As Boolean
    Dim ifCode As String
    Dim FrameName As String
    Dim LastFrame As CFrame
    Dim srcFile As String
    
    If (Not (UsesProtocol(FileName) Or IsExternalLink(FileName))) Or Left(FileName, 2) = "\\" Then
    
        Exists = FileExists(FileName)
    
        If Not Exists Then
            FileName = GetFilePath(MainFile) + FileName
            Exists = FileExists(FileName)
        End If
        
        If Exists Then
            cCode = LoadFile(FileName)
            
            p1 = 0
            Do
                p1 = InStr(p1 + 1, cCode, "<iframe", vbTextCompare)
                p2 = InStr(p1 + 1, cCode, "</iframe>", vbTextCompare)
                If p1 <> 0 Then
                    ifCode = Mid(cCode, p1, p2 - p1 + Len("</iframe>"))
                    srcFile = ObtainAbsPath(FileName, GetParamVal(ifCode, "src"))
                    FrameName = GetParamVal(ifCode, "name")
                    If LenB(FrameName) = 0 And LenB(srcFile) <> 0 Then FrameName = "frames[" & UBound(Frames) & "]"
                    If LenB(FrameName) <> 0 Then AddNewFrame pFrame.Name + "." + FrameName, srcFile
                    
                    If LenB(srcFile) <> 0 Then
                        LastFrame = Frames(UBound(Frames))
                        StartParsing srcFile, LastFrame
                    End If
                End If
            Loop Until p1 = 0
        End If
    End If
    
    GetIFrames = (p2 <> 0)

End Function

Public Function ObtainAbsPath(ParentFile As String, ByVal srcFile As String) As String

    If FileExists(srcFile) And InStr(srcFile, "\") > 0 Then
        ObtainAbsPath = srcFile
    Else
        If Not (UsesProtocol(srcFile) Or IsExternalLink(srcFile)) Then
            ObtainAbsPath = Replace(SetSlashDir(GetFilePath(ParentFile) + srcFile, sdBack), "\\", "\")
            ObtainAbsPath = Replace(ObtainAbsPath, "\.\", "\")
        End If
    End If
    
    ObtainAbsPath = DecodeUrl(ObtainAbsPath)
    
End Function

Private Function GetFrameSet(ByVal FileName As String) As String()

    Dim cCode() As String
    Dim Exists As Boolean
    
    ReDim cCode(0)
    cCode(0) = 0
    If (Not (UsesProtocol(FileName) Or IsExternalLink(FileName))) Or Left(FileName, 2) = "\\" Then
    
        Exists = FileExists(FileName)
    
        If Not Exists Then
            FileName = GetFilePath(MainFile) + FileName
            Exists = FileExists(FileName)
        End If
        
        If Exists Then cCode = extractFramesets(LoadFile(FileName))
    End If
    
    GetFrameSet = cCode
    
End Function

Private Function extractFramesets(c As String) As String()

    Dim p1 As Long
    Dim p2 As Long
    Dim f() As String

    ReDim f(0)
    f(0) = 0
    p1 = 1
    Do
        p1 = InStr(p1, c, "<frameset ", vbTextCompare)
        p2 = InStr(p1 + 1, c, "</frameset>", vbTextCompare)
        If p1 = 0 Or p2 = 0 Then Exit Do
        ReDim Preserve f(UBound(f) + 1)
        f(UBound(f)) = Mid(c, p1, p2 - p1 + Len("</frameset>"))
        p1 = p2
    Loop
    
    f(0) = UBound(f)
    extractFramesets = f
    
End Function

Private Function fsparser(ByVal ParentFile As String, ByVal cCode As String, pFrame As CFrame)

    Dim p1 As Long
    Dim p2 As Long
    Dim FrameName As String
    Dim srcFile As String
    Dim LastFrame As CFrame
    
    Do
        p1 = FindFrame(cCode)
        If p1 = 0 Then Exit Do
        
        FrameName = GetParamVal(cCode, "name")
        srcFile = ObtainAbsPath(ParentFile, GetParamVal(cCode, "src"))
        If LenB(FrameName) = 0 And LenB(srcFile) <> 0 Then FrameName = UBound(Frames)
        If LenB(FrameName) <> 0 Then AddNewFrame pFrame.Name + "." + FrameName, srcFile
        
        LastFrame = Frames(UBound(Frames))
        If FileExists(srcFile) Then StartParsing srcFile, LastFrame
        cCode = Mid(cCode, InStr(p1, cCode, ">") + 1)
        
        p1 = InStr(LCase(cCode), "<frameset ")
        p2 = InStr(LCase(cCode), "<frame ")
        If p1 < p2 And p1 <> 0 Then
            If Mid(LCase(cCode), p1, 9) = "<frameset" Then
                Dim c() As String
                Dim i As Integer
                
                c = extractFramesets(cCode)
                For i = 1 To Val(c(0))
                    fsparser ParentFile, c(i), LastFrame
                Next i
                
                p1 = InStr(LCase(cCode), "</frameset>")
                cCode = Mid(cCode, p1 + Len("</frameset>"))
            End If
        End If
    Loop

End Function

'Private Function fsparser2(ByVal ParentFile As String, ByVal cCode As String, pFrame As CFrame, Optional p1 As Long = 0)
'
'    Dim p2 As Long
'    Dim FrameName As String
'    Dim srcFile As String
'    Dim LastFrame As CFrame
'
'    On Error Resume Next
'
'    'Get the actual frameset contents
'    If p1 = 0 Then p1 = FindFrame(cCode)
'    p2 = InStrRev(cCode, "</frameset>", -1, vbTextCompare)
'    If p1 = 0 Or p2 = 0 Or p1 > p2 Then Exit Function
'    cCode = Mid(cCode, p1, p2 - p1)
'
'    p1 = 1
'    Do While cCode <> ""
'        p1 = FindFrame(cCode)
'        p2 = InStr(1, cCode, "<frameset ", vbTextCompare)
'        If p1 = 0 Then Exit Do
'        If p2 < p1 And p2 <> 0 Then
'            fsparser ParentFile, cCode, pFrame, p1
'
'            p1 = InStr(1, cCode, "</frameset>", vbTextCompare)
'            cCode = Mid(cCode, p1 + Len("</frameset>"))
'        Else
'            FrameName = GetParamVal(cCode, "name")
'            srcFile = ObtainAbsPath(ParentFile, GetParamVal(cCode, "src"))
'            If FrameName = "" And srcFile <> "" Then FrameName = UBound(Frames)
'            If FrameName <> "" Then AddNewFrame pFrame.Name + "." + FrameName, srcFile
'
'            LastFrame = Frames(UBound(Frames))
'            StartParsing srcFile, LastFrame
'            cCode = Mid(cCode, InStr(cCode, ">") + 1)
'        End If
'    Loop
'
'    If NagScreenIsVisible Then DoEvents
'
'End Function

Private Function FindFrame(ByVal c As String) As Integer

    Dim i As Long
    
    c = LCase(c)
    Do
        i = InStr(i + 1, c, "<frame")
    Loop While IsAlphaNum(Mid(c, i + 6, 1)) And (i > 0)
    FindFrame = i
    
End Function

Private Sub AddNewFrame(FramePath As String, SourceFile As String)

    Dim i As Integer
    Dim t As Integer
    
    t = UBound(Frames)
    For i = 1 To t
        If Frames(i).Name = FramePath Then Exit Sub
    Next i

    t = t + 1
    ReDim Preserve Frames(t)
    With Frames(t)
        .Name = FramePath
        .srcFile = SourceFile
    End With
    
End Sub

Private Function IsAlphaNum(ByVal k As String) As Boolean

    k = LCase(k)
    IsAlphaNum = (k >= "0" And k <= "9") Or (k >= "a" And k <= "z")

End Function

Public Function GetParamVal(c As String, ByVal p As String) As String

    Dim p1 As Long
    Dim p2 As Long
    Dim i As Long
    Dim k As String
    Dim ec As String
    
    i = 1
ReStart:
    i = InStr(i, c, p, vbTextCompare)
    If i > 0 Then
        If InStr(i, c, p, vbTextCompare) > 0 Then
        
            'Find "="
            For i = i + Len(p) To Len(c)
                k = Mid(c, i, 1)
                If k = "=" Then Exit For
                If IsAlphaNum(k) Then
                    GoTo ReStart
                End If
            Next i
            
            'Find first char of param
            For i = i + 1 To Len(c)
                k = Mid(c, i, 1)
                If k <> " " And k <> """" And k <> "'" And k <> vbTab And k <> vbCr And k <> vbLf Then
                    Exit For
                End If
                'Get the enclosing char
                If k <> " " Then ec = k
            Next i
            
            'Get the closing char for the param
            If LenB(ec) = 0 Then
                p1 = InStr(i, c, " ")
                p2 = InStr(i, c, ">")
                If p1 = 0 Then p1 = p2
                If p2 <> 0 And p2 < p1 Then p1 = p2
            Else
                p1 = InStr(i, c, ec)
            End If
            
            If p1 > 0 Then
                GetParamVal = Mid(c, i, p1 - i)
            End If
        End If
    End If
    
End Function

Public Function ChangeParamVal(c As String, ByVal p As String, ByVal n As String, Optional AllowAdd As Boolean = False) As String

    Dim p1 As Long
    Dim p2 As Long
    Dim i As Long
    Dim k As String
    Dim ec As String
    Dim ns As String
    
    i = 1
ReStart:
    i = InStr(i, c, p, vbTextCompare)
    If i > 0 Then
        If InStr(i, c, p, vbTextCompare) > 0 Then
        
            'Find "="
            For i = i + Len(p) To Len(c)
                k = Mid(c, i, 1)
                If k = "=" Then Exit For
                If IsAlphaNum(k) Then
                    GoTo ReStart
                End If
            Next i
            
            'Find first char of param
            For i = i + 1 To Len(c)
                k = Mid(c, i, 1)
                If k <> " " And k <> """" And k <> "'" And k <> vbTab And k <> vbCr And k <> vbLf Then
                    Exit For
                End If
                'Get the enclosing char
                If k <> " " Then ec = k
            Next i
            
            ns = Left(c, i - 1) + n
            
            'Get the closing char for the param
            If LenB(ec) = 0 Then
                p1 = InStr(i, c, " ")
                p2 = InStr(i, c, ">")
                If p1 = 0 Then p1 = p2
                If p2 <> 0 And p2 < p1 Then p1 = p2
            Else
                p1 = InStr(i, c, ec)
            End If
            
            ns = ns + Mid(c, p1)
        End If
    Else
        If AllowAdd Then
            i = InStr(c, ">")
            ns = Left(c, i - 1) + " " + p + "=""" + n + """" + Mid(c, i)
        End If
    End If
    
    ChangeParamVal = IIf(LenB(ns) = 0, c, ns)
    
End Function
