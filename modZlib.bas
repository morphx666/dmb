Attribute VB_Name = "modZlib"
Option Explicit

Public Sub Compress(ByVal FileName As String, ByVal sCode As String, Optional GZIPcompat As Boolean = False)

    On Error Resume Next

    CreateCtrl
    
    SaveFile FileName, sCode

    With frmMain.Controls("zlibCtrl")
        .InputFile = FileName
        .OutputFile = FileName + ".tmp"
        If FileExists(.OutputFile) Then Kill .OutputFile
        .level = 9 'Maximum
        .Compress
        Kill .InputFile
        Name CStr(.OutputFile) As CStr(.InputFile)
        
        If GZIPcompat Then
            Dim sH As String
            sCode = LoadFile(.InputFile)
            sH = sH + Chr(&H1F)                 ' 00 - GZIP Magic
            sH = sH + Chr(&H8B)                 ' 01 - GZIP Magic
            sH = sH + Chr(&H8)                  ' 02 - Compression Level
            sH = sH + Chr(&H0)                  ' 03 - Flags (none)
            sH = sH + Chr(&H0)                  ' 04 - Time
            sH = sH + Chr(&H0)                  ' 05 - Time
            sH = sH + Chr(&H0)                  ' 06 - Time
            sH = sH + Chr(&H0)                  ' 07 - Time
            sH = sH + Chr(&H2)                  ' 08 - Additional compression flags (compressor used maximum compression)
            sH = sH + Chr(&HB)                  ' 09 - Operating System (NTFS FileSystem)
            
            SaveFile .InputFile + ".gz", sH + Mid(sCode, 3)
            Kill .InputFile
        End If
    End With
    
    DestroyCtrl

End Sub

Public Function UnCompress(ByVal FileName As String) As String

    On Error Resume Next
    
    CreateCtrl

    With frmMain.Controls("zlibCtrl")
        .InputFile = FileName
        .OutputFile = FileName + ".tmp"
        If FileExists(.OutputFile) Then Kill .OutputFile
        .Decompress
        UnCompress = LoadFile(.OutputFile)
        Kill .OutputFile
    End With
    
    DestroyCtrl

End Function

Private Sub CreateCtrl()

    On Error Resume Next

    DestroyCtrl
    
    frmMain.Controls.Add "ZLIBTOOL.ZlibToolCtrl.1", "zlibCtrl"

End Sub

Private Sub DestroyCtrl()

    On Error Resume Next
    
    frmMain.Controls.Remove "zlibCtrl"

End Sub
