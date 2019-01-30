Attribute VB_Name = "modImgHandling"
Option Explicit

Private AltProject As String

Public Function IsFlash(FileName As String) As Boolean

    On Error GoTo ExitSub
    
    IsFlash = Right(LCase(FileName), 4) = ".swf"
    
ExitSub:

End Function

'Public Function IsANI(FileName As String) As Boolean
'
'    On Error GoTo ExitSub
'
'    IsANI = Right(LCase(FileName), 4) = ".ani"
'
'ExitSub:
'
'End Function

#If STANDALONE = 0 Then

Public Sub SaveImages(UseAltProject As String)

    Dim i As Integer
    Dim ff As Integer
    
    AltProject = UseAltProject
    
    AddImageResource 0, "", True
    ff = FreeFile
    Open Project.FileName For Binary As #ff
    Seek #ff, LOF(ff) + 1
    Put #ff, , "[RSC]" + vbCrLf
    
    For i = 1 To UBound(Project.Toolbars)
        AddImageResource ff, Project.Toolbars(i).Image
    Next i
    
    For i = 1 To UBound(MenuGrps)
        AddImageResource ff, MenuGrps(i).Image
    
        AddImageResource ff, MenuGrps(i).LeftImage.NormalImage
        AddImageResource ff, MenuGrps(i).LeftImage.HoverImage
    
        AddImageResource ff, MenuGrps(i).RightImage.NormalImage
        AddImageResource ff, MenuGrps(i).RightImage.HoverImage
    
        AddImageResource ff, MenuGrps(i).CornersImages.gcTopLeft
        AddImageResource ff, MenuGrps(i).CornersImages.gcTopCenter
        AddImageResource ff, MenuGrps(i).CornersImages.gcTopRight
    
        AddImageResource ff, MenuGrps(i).CornersImages.gcLeft
        AddImageResource ff, MenuGrps(i).CornersImages.gcRight
    
        AddImageResource ff, MenuGrps(i).CornersImages.gcBottomLeft
        AddImageResource ff, MenuGrps(i).CornersImages.gcBottomCenter
        AddImageResource ff, MenuGrps(i).CornersImages.gcBottomRight
    
        AddImageResource ff, MenuGrps(i).iCursor.CFile
    Next i
    
    For i = 1 To UBound(MenuCmds)
        AddImageResource ff, MenuCmds(i).BackImage.NormalImage
        AddImageResource ff, MenuCmds(i).BackImage.HoverImage
    
        AddImageResource ff, MenuCmds(i).LeftImage.NormalImage
        AddImageResource ff, MenuCmds(i).LeftImage.HoverImage
    
        AddImageResource ff, MenuCmds(i).RightImage.NormalImage
        AddImageResource ff, MenuCmds(i).RightImage.HoverImage
    
        AddImageResource ff, MenuCmds(i).iCursor.CFile
    Next i
    
    Close #ff

End Sub

Private Sub AddImageResource(ff As Integer, FileName As String, Optional Reset As Boolean)

    On Error Resume Next
    
    Dim i As Integer
    Static SavedImages() As String
    
    If Reset Then
        ReDim SavedImages(0)
        Exit Sub
    End If
    
    If FileName = "" Then Exit Sub
    
    For i = 1 To UBound(SavedImages)
        If SavedImages(i) = FileName Then Exit Sub
    Next i
    ReDim Preserve SavedImages(UBound(SavedImages) + 1)
    SavedImages(UBound(SavedImages)) = FileName

    Put #ff, , "RSCImg::" & FileName & "::"
    
    If FileExists(FileName) Then
        Put #ff, , GetImgCode_FromFile(FileName)
    Else
        Put #ff, , GetImgCode_FromRes(FileName)
    End If
    Put #ff, , vbCrLf

End Sub

Private Function GetImgSize(FileName As String) As Long

'    Dim ff As Long
'
'    ff = FreeFile
'    Open FileName For Binary As #ff
'        GetImgSize = LOF(ff)
'    Close #ff

    GetImgSize = FileLen(FileName)

End Function

Private Function GetImgCode_FromFile(FileName As String) As Byte()

'    Dim ff As Long
'    Dim imgCode() As Byte
'
'    ff = FreeFile
'    Open FileName For Binary As #ff
'        ReDim imgCode(LOF(ff) - 1)
'        Get #ff, , imgCode
'    Close #ff
'
'    GetImgCode_FromFile = imgCode

    GetImgCode_FromFile = LoadFile(FileName)

End Function

Public Function CalcImagesSize(Images() As String) As Long

    Dim c As Integer
    Dim g As Integer
    Dim j As Integer
    Dim Size As Long
    
    ReDim Images(0)
    
    For c = 1 To UBound(Project.Toolbars)
        If Project.Toolbars(c).Image <> "" Then
            ReDim Preserve Images(UBound(Images) + 1)
            Images(UBound(Images)) = Project.Toolbars(c).Image
        End If
    Next c
    
    For c = 1 To UBound(MenuCmds)
        With MenuCmds(c).LeftImage
            If .NormalImage <> "" Then
                ReDim Preserve Images(UBound(Images) + 2)
                Images(UBound(Images) - 1) = .NormalImage
                Images(UBound(Images)) = .HoverImage
            End If
        End With
        With MenuCmds(c).RightImage
            If .NormalImage <> "" Then
                ReDim Preserve Images(UBound(Images) + 2)
                Images(UBound(Images) - 1) = .NormalImage
                Images(UBound(Images)) = .HoverImage
            End If
        End With
        With MenuCmds(c).BackImage
            If .NormalImage <> "" Then
                ReDim Preserve Images(UBound(Images) + 2)
                Images(UBound(Images) - 1) = .NormalImage
                Images(UBound(Images)) = .HoverImage
            End If
        End With
    Next c
    For g = 1 To UBound(MenuGrps)
        With MenuGrps(g)
            If .Image <> "" Then
                ReDim Preserve Images(UBound(Images) + 1)
                Images(UBound(Images)) = .Image
            End If
            With .BackImage
                If .NormalImage <> "" Then
                    ReDim Preserve Images(UBound(Images) + 2)
                    Images(UBound(Images) - 1) = .NormalImage
                    Images(UBound(Images)) = .HoverImage
                End If
            End With
            With .LeftImage
                If .NormalImage <> "" Then
                    ReDim Preserve Images(UBound(Images) + 2)
                    Images(UBound(Images) - 1) = .NormalImage
                    Images(UBound(Images)) = .HoverImage
                End If
            End With
            With .RightImage
                If .NormalImage <> "" Then
                    ReDim Preserve Images(UBound(Images) + 2)
                    Images(UBound(Images) - 1) = .NormalImage
                    Images(UBound(Images)) = .HoverImage
                End If
            End With
        End With
    Next g
    
ReStart:
    For c = 1 To UBound(Images)
        If Images(c) = "" Then
            For g = c To UBound(Images) - 1
                Images(g) = Images(g + 1)
            Next g
            ReDim Preserve Images(UBound(Images) - 1)
            GoTo ReStart
        End If
        For g = c + 1 To UBound(Images)
            If Images(c) = Images(g) Then
                For j = c To UBound(Images) - 1
                    Images(j) = Images(j + 1)
                Next j
                ReDim Preserve Images(UBound(Images) - 1)
                GoTo ReStart
            End If
        Next g
    Next c
    
    For j = 1 To UBound(Images)
        Size = Size + GetImgSize(TempPath + GetFileName(Images(j)))
    Next j
    
    CalcImagesSize = Size

End Function

Private Sub RemoveFromArray(a() As String, i As Integer)

    Dim j As Integer

    For j = i To UBound(a) - 1
        a(j) = a(j + 1)
    Next j
    ReDim Preserve a(UBound(a) - 1)

End Sub

#End If

Public Function LoadPictureRes(ByVal FileName As String) As IPictureDisp

    Dim imgCode As String
    
    On Error Resume Next
    
    If FileName <> "" Then
        If Not FileExists(FileName) Then
            imgCode = LoadFile(Project.FileName)
            imgCode = Mid(imgCode, InStr(imgCode, FileName + "::") + Len(FileName + "::"))
            If InStr(imgCode, "RSCImg::") > 0 Then imgCode = Left(imgCode, InStr(imgCode, "RSCImg::") - 3)
            
            FileName = StatesPath + GetFileName(FileName)
            
            SaveFile FileName, imgCode
        End If
        
        Set LoadPictureRes = OLLoadPicture(FileName)
        'Err.Clear
        'Set LoadPictureRes = LoadPicture(FileName, , vbLPColor)
        'If Err.Number Then
        '    Set LoadPictureRes = LoadPicture()
        '    'MsgBox "The " + GetFileName(FileName) + " could not be loaded.", vbInformation + vbOKOnly, "Invalid Image"
        'End If
    Else
        Set LoadPictureRes = LoadPicture()
    End If

End Function

Public Function OLLoadPicture(FileName As String) As IPictureDisp

    Static GflAxObj As Object
    
    If FileName = "" Then
        Set OLLoadPicture = LoadPicture()
    Else
        On Error Resume Next
        Set OLLoadPicture = LoadPicture(FileName)
        If Err.Number Then
            Err.Clear
            If GflAxObj Is Nothing Then
                Set GflAxObj = CreateObject("GflAx170.GflAx")
                GflAxObj.EnableLZW = True
            End If
            With GflAxObj
                .LoadBitmap FileName
                Set OLLoadPicture = .GetPicture
            End With
        End If
    End If

End Function

Public Sub CopyProjectImages(trgPath As String)

    Dim c As Integer
    Dim g As Integer
    
    'This function also copies the sounds
    'The On Error trap is used to capture errors on sounds
    On Error Resume Next
    
    For c = 1 To UBound(Project.Toolbars)
        CopyRes2File Project.Toolbars(c).Image, trgPath + GetFileName(Project.Toolbars(c).Image)
    Next c

    For c = 1 To UBound(MenuCmds)
        With MenuCmds(c)
            With .LeftImage
                CopyRes2File .NormalImage, trgPath + GetFileName(.NormalImage)
                CopyRes2File .HoverImage, trgPath + GetFileName(.HoverImage)
            End With
            With .RightImage
                CopyRes2File .NormalImage, trgPath + GetFileName(.NormalImage)
                CopyRes2File .HoverImage, trgPath + GetFileName(.HoverImage)
            End With
            With .BackImage
                CopyRes2File .NormalImage, trgPath + GetFileName(.NormalImage)
                CopyRes2File .HoverImage, trgPath + GetFileName(.HoverImage)
            End With
            'With .Sound
            '    If .onmouseover <> "" Then FileCopy .onmouseover, trgPath + GetFileName(.onmouseover)
            '    If .onclick <> "" Then FileCopy .onclick, trgPath + GetFileName(.onclick)
            'End With
            With .iCursor
                If .cType = iccCustom Then FileCopy .CFile, trgPath + GetFileName(.CFile)
            End With
        End With
    Next c
    For g = 1 To UBound(MenuGrps)
        With MenuGrps(g)
            CopyRes2File .Image, trgPath + GetFileName(.Image)
            With .BackImage
                CopyRes2File .NormalImage, trgPath + GetFileName(.NormalImage)
                CopyRes2File .HoverImage, trgPath + GetFileName(.HoverImage)
            End With
            With .LeftImage
                CopyRes2File .NormalImage, trgPath + GetFileName(.NormalImage)
                CopyRes2File .HoverImage, trgPath + GetFileName(.HoverImage)
            End With
            With .RightImage
                CopyRes2File .NormalImage, trgPath + GetFileName(.NormalImage)
                CopyRes2File .HoverImage, trgPath + GetFileName(.HoverImage)
            End With
            'With .Sound
            '    If .onmouseover <> "" Then FileCopy .onmouseover, trgPath + GetFileName(.onmouseover)
            '    If .onclick <> "" Then FileCopy .onclick, trgPath + GetFileName(.onclick)
            'End With
            
            With .CornersImages
                CopyRes2File .gcTopLeft, trgPath + GetFileName(.gcTopLeft)
                CopyRes2File .gcTopCenter, trgPath + GetFileName(.gcTopCenter)
                CopyRes2File .gcTopRight, trgPath + GetFileName(.gcTopRight)
                
                CopyRes2File .gcLeft, trgPath + GetFileName(.gcLeft)
                CopyRes2File .gcRight, trgPath + GetFileName(.gcRight)
                
                CopyRes2File .gcBottomLeft, trgPath + GetFileName(.gcBottomLeft)
                CopyRes2File .gcBottomCenter, trgPath + GetFileName(.gcBottomCenter)
                CopyRes2File .gcBottomRight, trgPath + GetFileName(.gcBottomRight)
            End With
            
            With .iCursor
                If .cType = iccCustom Then FileCopy .CFile, trgPath + GetFileName(.CFile)
            End With
        End With
    Next g
    
    On Error Resume Next
    
    'If Project.UnfoldingSound.onmouseover <> "" Then
    '    FileCopy Project.UnfoldingSound.onmouseover, trgPath + GetFileName(Project.UnfoldingSound.onmouseover)
    'End If

End Sub

Private Sub CopyRes2File(srcFileName As String, trgFileName As String)

    If srcFileName = "" Then Exit Sub
    SaveFile trgFileName, GetImgCode_FromRes(srcFileName)

End Sub

Public Function GetImgCode_FromRes(ByVal FileName As String) As String

    Dim imgCode As String
    Dim cStart As Long

    If FileExists(FileName) Then
        imgCode = LoadFile(FileName)
    Else
        If FileExists(AltProject) Then
            imgCode = LoadFile(AltProject)
        Else
            imgCode = LoadFile(Project.FileName)
        End If
        If InStr(imgCode, FileName + "::") = 0 Then Exit Function
        imgCode = Mid(imgCode, InStr(imgCode, FileName + "::") + Len(FileName + "::"))
        If InStr(imgCode, "RSCImg::") > 0 Then imgCode = Left(imgCode, InStr(imgCode, "RSCImg::") - 3)
        
        FileName = StatesPath + GetFileName(FileName)
        
        SaveFile FileName, imgCode
    End If
    
    GetImgCode_FromRes = imgCode

End Function
