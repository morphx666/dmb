Attribute VB_Name = "modGetIcon"
Option Explicit

Public Const DI_MASK = &H1
Public Const DI_IMAGE = &H2
Public Const DI_NORMAL = DI_MASK Or DI_IMAGE

Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hDc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long

Private Type SHFILEINFO
    hIcon As Long                      '  out: icon
    iIcon As Long                      '  out: icon index
    dwAttributes As Long               '  out: SFGAO_ flags
    szDisplayName As String * 260      '  out: display name (or path)
    szTypeName As String * 80          '  out: type name
End Type

Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long

Private Const SHGFI_ICON = &H100
'Private Const SHGFI_LARGEICON = &H0
Private Const SHGFI_SMALLICON = &H1
'Private Const SHGFI_DISPLAYNAME = &H200

Public Sub GetIcon(tPic As PictureBox, fName As String, Optional s As Long)
    
    Dim sh_info As SHFILEINFO
    
    tPic.Picture = LoadPicture()
    
    If s = 0 Then s = 16
    SHGetFileInfo fName, 0, sh_info, Len(sh_info), SHGFI_ICON + SHGFI_SMALLICON
    DrawIconEx tPic.hDc, 0, 0, sh_info.hIcon, s, s, 0, 0, DI_NORMAL
    
    DestroyIcon sh_info.hIcon
    ReleaseDC tPic.Parent.hWnd, tPic.hDc

End Sub

