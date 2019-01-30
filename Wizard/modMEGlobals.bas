Attribute VB_Name = "modMEGlobals"
Option Explicit

Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    ' Optional
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type

Private Const SEE_MASK_INVOKEIDLIST = &HC
Private Const SEE_MASK_NOCLOSEPROCESS = &H40
Private Const SEE_MASK_FLAG_NO_UI = &H400
Private Declare Function ShellExecuteEx Lib "shell32.dll" (lps As SHELLEXECUTEINFO) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long

Public Function RunShellExecuteLink(sTopic As String, sFile As Variant, sParams As Variant, sDirectory As Variant, nShowCmd As Long) As Long

    Dim hWndDesk As Long
    Dim lps As SHELLEXECUTEINFO
    
    With lps
        .hwnd = GetDesktopWindow()
        .fMask = SEE_MASK_NOCLOSEPROCESS
        .lpVerb = sTopic
        .nShow = nShowCmd
        .lpFile = sFile
        .lpDirectory = sDirectory
        .cbSize = Len(lps)
    End With
    
    ShellExecuteEx lps
    
    RunShellExecuteLink = lps.hProcess

End Function

