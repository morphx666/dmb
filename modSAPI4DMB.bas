Attribute VB_Name = "modSAPI4DMB"
Option Explicit

Public Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" _
    (ByVal lpRootPathName As String, _
    ByVal lpVolumeNameBuffer As String, _
    ByVal nVolumeNameSize As Long, _
    lpVolumeSerialNumber As Long, _
    lpMaximumComponentLength As Long, _
    lpFileSystemFlags As Long, _
    ByVal lpFileSystemNameBuffer As String, _
    ByVal nFileSystemNameSize As Long) _
    As Long
    
Public Function GetHDSerial(Optional ByVal DoMD5 As Boolean = True) As String

    Dim volName As String
    Dim drvSerial As String
    Dim clsMD5 As MD5
    
    Set clsMD5 = New MD5
    rgbGetVolume "c:\", volName, drvSerial
    
    If DoMD5 Then
        GetHDSerial = clsMD5.DigestStrToHexStr(drvSerial)
    Else
        GetHDSerial = drvSerial
    End If

End Function

Private Sub rgbGetVolume(PathName As String, DrvVolumeName As String, DrvSerialNo As String)
 
  'create working variables
  'to keep it simple, use dummy variables for info
  'we're not interested in right now
   Dim r As Long
   Dim pos As Integer
   Dim HiWord As Long
   Dim HiHexStr As String
   Dim LoWord As Long
   Dim LoHexStr As String
   Dim VolumeSN As Long

   Dim UnusedStr As String
   Dim UnusedVal1 As Long
   Dim UnusedVal2 As Long

  'pad the strings
   DrvVolumeName$ = Space$(14)
   UnusedStr$ = Space$(32)

  'do what it says
   r = GetVolumeInformation(PathName, _
                            DrvVolumeName, _
                            Len(DrvVolumeName), _
                            VolumeSN&, _
                            UnusedVal1, UnusedVal2, _
                            UnusedStr, Len(UnusedStr$))


  'error check
   If r& = 0 Then Exit Sub

  'determine the volume label
   pos = InStr(DrvVolumeName, ChrW$(0))
   If pos Then DrvVolumeName = Left$(DrvVolumeName, pos - 1)
   If Len(Trim$(DrvVolumeName)) = 0 Then DrvVolumeName = "(no label)"

  'determine the drive volume id
   HiWord = GetHiWord(VolumeSN) And &HFFFF&
   LoWord = GetLoWord(VolumeSN) And &HFFFF&
   HiHexStr = PadZeros(Hex(HiWord), 4)
   LoHexStr = PadZeros(Hex(LoWord), 4)
 
   DrvSerialNo = HiHexStr & "-" & LoHexStr

End Sub

Private Function PadZeros(Num As String, Size As Integer) As String

    If Size < Len(Num) Then
        'Debug.Print "Padding Error"
    Else
        PadZeros = String(Size - Len(Num), "0") & Num
    End If

End Function

Private Function GetHiWord(dw As Long) As Integer
  
    If dw And &H80000000 Then
          GetHiWord = (dw \ 65535) - 1
    Else: GetHiWord = dw \ 65535
    End If
    
End Function
  
Private Function GetLoWord(dw As Long) As Integer
  
    If dw And &H8000& Then
          GetLoWord = &H8000 Or (dw And &H7FFF&)
    Else: GetLoWord = dw And &HFFFF&
    End If
    
End Function
