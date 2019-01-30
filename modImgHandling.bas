Attribute VB_Name = "modImgHandling"
Option Explicit

Private AltProject As String

Private Declare Function LoadFileAsHEX Lib "xfxbinimg.dll" (ByVal FileName As String, ByVal fCode As String) As Long
Private Declare Function SaveFileFromHEX Lib "xfxbinimg.dll" (ByVal FileName As String, ByVal fCode As String, ByVal Size As Long) As Long

Public Function IsFlash(FileName As String) As Boolean

    On Error GoTo ExitSub
    
    IsFlash = LCase(GetFileExtension(FileName)) = "swf"
    
ExitSub:

End Function

Public Function IsANI(FileName As String) As Boolean

    On Error GoTo ExitSub

    IsANI = LCase(GetFileExtension(FileName)) = "ani"

ExitSub:

End Function

#If STANDALONE = 0 Then

Public Sub SaveImages(UseAltProject As String)

    Dim i As Integer
    Dim ff As Integer
    
    AltProject = UseAltProject
    
    AddImageResource 0, vbNullString, True
    ff = FreeFile
    Open Project.FileName For Binary As #ff
    Seek #ff, LOF(ff) + 1
    Put #ff, , "[RSC]" + vbCrLf
    
    For i = 1 To UBound(Project.Toolbars)
        AddImageResource ff, Project.Toolbars(i).Image
    Next i
    
    For i = 1 To UBound(MenuGrps)
        With MenuGrps(i)
            If .Compile Then
                AddImageResource ff, .Image
            
                AddImageResource ff, .tbiLeftImage.NormalImage
                AddImageResource ff, .tbiLeftImage.HoverImage
            
                AddImageResource ff, .tbiRightImage.NormalImage
                AddImageResource ff, .tbiRightImage.HoverImage
                
                AddImageResource ff, .tbiBackImage.NormalImage
                AddImageResource ff, .tbiBackImage.HoverImage
            
                AddImageResource ff, .CornersImages.gcTopLeft
                AddImageResource ff, .CornersImages.gcTopCenter
                AddImageResource ff, .CornersImages.gcTopRight
            
                AddImageResource ff, .CornersImages.gcLeft
                AddImageResource ff, .CornersImages.gcRight
            
                AddImageResource ff, .CornersImages.gcBottomLeft
                AddImageResource ff, .CornersImages.gcBottomCenter
                AddImageResource ff, .CornersImages.gcBottomRight
                
                AddImageResource ff, .scrolling.DnImage.NormalImage
                AddImageResource ff, .scrolling.DnImage.HoverImage
                AddImageResource ff, .scrolling.UpImage.NormalImage
                AddImageResource ff, .scrolling.UpImage.HoverImage
            
                AddImageResource ff, .iCursor.CFile
            End If
        End With
    Next i
    
    For i = 1 To UBound(MenuCmds)
        With MenuCmds(i)
            If .Compile Then
                AddImageResource ff, .BackImage.NormalImage
                AddImageResource ff, .BackImage.HoverImage
            
                AddImageResource ff, .LeftImage.NormalImage
                AddImageResource ff, .LeftImage.HoverImage
            
                AddImageResource ff, .RightImage.NormalImage
                AddImageResource ff, .RightImage.HoverImage
            
                AddImageResource ff, .iCursor.CFile
            End If
        End With
    Next i
    
    Close #ff

End Sub

Private Sub AddImageResource(ff As Integer, FileName As String, Optional Reset As Boolean)

    On Error Resume Next
    
    'Dim t As Single
    'Static cc As Single
    't = Timer
    
    Dim i As Integer
    Dim sCode As String
    Static SavedImages() As String
    
    If Reset Then
        'cc = 0
        ReDim SavedImages(0)
        Exit Sub
    End If
    
    If LenB(FileName) = 0 Then Exit Sub
    
    For i = 1 To UBound(SavedImages)
        If SavedImages(i) = FileName Then Exit Sub
    Next i
    ReDim Preserve SavedImages(UBound(SavedImages) + 1)
    SavedImages(UBound(SavedImages)) = FileName

    Put #ff, , "RSCImg::" & FileName & "::"
    
    If FileExists(FileName) Then
        sCode = GetImgCode_FromFile(FileName)
    Else
        sCode = GetImgCode_FromRes(FileName)
    End If
    Put #ff, , sCode
    Put #ff, , vbCrLf
    
    't = Round(Timer - t, 2)
    'cc = cc + t
    'Debug.Print GetFileName(FileName) + ": " & t & "( " & cc & ")"

End Sub

Private Function GetImgSize(FileName As String) As Long

    GetImgSize = FileLen(FileName)

End Function

Private Function GetImgCode_FromFile(FileName As String) As Byte()
    
    GetImgCode_FromFile = LoadImageFile(FileName)

End Function

Public Function CalcImagesSize(Images() As String) As Long

    Dim c As Integer
    Dim g As Integer
    Dim j As Integer
    Dim Size As Long
    
    ReDim Images(0)
    
    For c = 1 To UBound(Project.Toolbars)
        If LenB(Project.Toolbars(c).Image) <> 0 Then
            ReDim Preserve Images(UBound(Images) + 1)
            Images(UBound(Images)) = Project.Toolbars(c).Image
        End If
    Next c
    
    For c = 1 To UBound(MenuCmds)
        With MenuCmds(c).LeftImage
            If LenB(.NormalImage) <> 0 Then
                ReDim Preserve Images(UBound(Images) + 2)
                Images(UBound(Images) - 1) = .NormalImage
                Images(UBound(Images)) = .HoverImage
            End If
        End With
        With MenuCmds(c).RightImage
            If LenB(.NormalImage) <> 0 Then
                ReDim Preserve Images(UBound(Images) + 2)
                Images(UBound(Images) - 1) = .NormalImage
                Images(UBound(Images)) = .HoverImage
            End If
        End With
        With MenuCmds(c).BackImage
            If LenB(.NormalImage) <> 0 Then
                ReDim Preserve Images(UBound(Images) + 2)
                Images(UBound(Images) - 1) = .NormalImage
                Images(UBound(Images)) = .HoverImage
            End If
        End With
    Next c
    For g = 1 To UBound(MenuGrps)
        With MenuGrps(g)
            If LenB(.Image) <> 0 Then
                ReDim Preserve Images(UBound(Images) + 1)
                Images(UBound(Images)) = .Image
            End If
            With .BackImage
                If LenB(.NormalImage) <> 0 Then
                    ReDim Preserve Images(UBound(Images) + 2)
                    Images(UBound(Images) - 1) = .NormalImage
                    Images(UBound(Images)) = .HoverImage
                End If
            End With
            With .tbiLeftImage
                If LenB(.NormalImage) <> 0 Then
                    ReDim Preserve Images(UBound(Images) + 2)
                    Images(UBound(Images) - 1) = .NormalImage
                    Images(UBound(Images)) = .HoverImage
                End If
            End With
            With .tbiRightImage
                If LenB(.NormalImage) <> 0 Then
                    ReDim Preserve Images(UBound(Images) + 2)
                    Images(UBound(Images) - 1) = .NormalImage
                    Images(UBound(Images)) = .HoverImage
                End If
            End With
            With .tbiBackImage
                If LenB(.NormalImage) <> 0 Then
                    ReDim Preserve Images(UBound(Images) + 2)
                    Images(UBound(Images) - 1) = .NormalImage
                    Images(UBound(Images)) = .HoverImage
                End If
            End With
        End With
    Next g
    
ReStart:
    For c = 1 To UBound(Images)
        If LenB(Images(c)) = 0 Then
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

#End If

Public Function LoadPictureRes(ByVal FileName As String) As IPictureDisp

    Dim imgCode As String
    
    On Error Resume Next
    
    If LenB(FileName) <> 0 Then
        If Not FileExists(FileName) Then
            imgCode = LoadFile(Project.FileName)
            imgCode = Mid(imgCode, InStr(1, imgCode, FileName + "::", vbTextCompare) + Len(FileName + "::"))
            If InStr(imgCode, "RSCImg::") > 0 Then imgCode = Left(imgCode, InStr(imgCode, "RSCImg::") - 3)
            
            FileName = StatesPath + GetFileName(FileName)
            
            SaveImageFile FileName, imgCode
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

    On Error Resume Next

    Static GflAxObj As Object
    
    If LenB(FileName) = 0 Then
        Set OLLoadPicture = LoadPicture()
    Else
        On Error Resume Next
        Set OLLoadPicture = LoadPicture(FileName)
        If Err.number Then
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

Public Sub CopyProjectImages(trgPath As String, Optional gidx As Integer = -1, Optional cidx As Integer = -1, Optional tidx As Integer = -1)

    Dim c As Integer
    Dim ct As Integer
    Dim g As Integer
    Dim gt As Integer
    Dim tt As Integer
    
    'This function also copies the sounds
    'The On Error trap is used to capture errors on sounds
    On Error Resume Next
    
    If cidx = -1 Then
        ct = UBound(MenuCmds)
    Else
        ct = IIf(cidx = -2, -3, cidx)
    End If
    If gidx = -1 Then
        gt = UBound(MenuGrps)
    Else
        gt = IIf(gidx = -2, -3, gidx)
    End If
    If tidx = -1 Then
        tt = UBound(Project.Toolbars)
    Else
        tt = IIf(tidx = -2, -3, tidx)
    End If
    
    With Project
        For c = tidx To tt
            CopyRes2File .Toolbars(c).Image, trgPath + GetFileName(.Toolbars(c).Image)
        Next c
        With .AutoScroll
            If .MaxHeight <> 0 Then
                CopyRes2File .DnImage.NormalImage, trgPath + GetFileName(.DnImage.NormalImage)
                CopyRes2File .DnImage.HoverImage, trgPath + GetFileName(.DnImage.HoverImage)
                CopyRes2File .UpImage.NormalImage, trgPath + GetFileName(.UpImage.NormalImage)
                CopyRes2File .UpImage.HoverImage, trgPath + GetFileName(.UpImage.HoverImage)
            End If
        End With
    End With

    For c = cidx To ct
        With MenuCmds(c)
            If gidx = -1 Or (MenuCmds(c).Parent = gidx) Then
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
            End If
        End With
        
        'DoEvents
    Next c
    For g = gidx To gt
        With MenuGrps(g)
            CopyRes2File .Image, trgPath + GetFileName(.Image)
            With .BackImage
                CopyRes2File .NormalImage, trgPath + GetFileName(.NormalImage)
                CopyRes2File .HoverImage, trgPath + GetFileName(.HoverImage)
            End With
            With .tbiLeftImage
                CopyRes2File .NormalImage, trgPath + GetFileName(.NormalImage)
                CopyRes2File .HoverImage, trgPath + GetFileName(.HoverImage)
            End With
            With .tbiRightImage
                CopyRes2File .NormalImage, trgPath + GetFileName(.NormalImage)
                CopyRes2File .HoverImage, trgPath + GetFileName(.HoverImage)
            End With
            With .tbiBackImage
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
            
            With .scrolling
                If .MaxHeight <> 0 Then
                    CopyRes2File .DnImage.NormalImage, trgPath + GetFileName(.DnImage.NormalImage)
                    CopyRes2File .DnImage.HoverImage, trgPath + GetFileName(.DnImage.HoverImage)
                    CopyRes2File .UpImage.NormalImage, trgPath + GetFileName(.UpImage.NormalImage)
                    CopyRes2File .UpImage.HoverImage, trgPath + GetFileName(.UpImage.HoverImage)
                End If
            End With
            
            With .iCursor
                If .cType = iccCustom Then
                    CopyRes2File .CFile, trgPath + GetFileName(.CFile)
                End If
            End With
        End With
        
        'DoEvents
    Next g

End Sub

Public Sub CopyRes2File(srcFileName As String, trgFileName As String)

    If LenB(srcFileName) = 0 Then Exit Sub
    SaveImageFile trgFileName, GetImgCode_FromRes(srcFileName)

End Sub

Public Function GetImgCode_FromRes(ByVal FileName As String) As String

    Dim imgCode As String
    Dim p1 As Long
    Dim p2 As Long

    If FileExists(FileName) Then
        imgCode = LoadImageFile(FileName)
    Else
        If FileExists(AltProject) Then
            imgCode = LoadFile(AltProject)
        Else
            imgCode = LoadFile(Project.FileName)
        End If
        p1 = InStr(imgCode, FileName + "::")
        If p1 = 0 Then Exit Function
        p1 = p1 + Len(FileName + "::")
        p2 = InStr(p1, imgCode, "RSCImg::"): If p2 = 0 Then p2 = Len(imgCode) + 1
        imgCode = Mid(imgCode, p1, p2 - p1 - 2)
        
        If Left(imgCode, 3) <> "HEX" Then
            SaveFile StatesPath + GetFileName(FileName), imgCode
            imgCode = LoadImageFile(StatesPath + GetFileName(FileName))
            Kill StatesPath + GetFileName(FileName)
        End If
    End If
    
    GetImgCode_FromRes = imgCode

End Function

Public Function LoadImageFile(FileName As String, Optional JustCode As Boolean = False) As String

    LoadImageFile = Space(2 * FileLen(FileName))
    LoadFileAsHEX FileName, LoadImageFile
    
    LoadImageFile = IIf(JustCode, vbNullString, "HEX") + LoadImageFile

End Function

Public Function SaveImageFile(FileName As String, ByVal imgCode As String) As String

    If Left(imgCode, 3) = "HEX" Then
        imgCode = Mid(imgCode, 4)
        SaveFileFromHEX FileName, imgCode, Len(imgCode)
    Else
        SaveFile FileName, imgCode
    End If
    
End Function
