VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Clean DE Samples"
   ClientHeight    =   2115
   ClientLeft      =   8625
   ClientTop       =   6120
   ClientWidth     =   4215
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrStart 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1650
      Top             =   420
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    tmrStart.Enabled = True
    Me.Left = -Me.Width * 2
    Me.Top = -Me.Height * 2

End Sub

Private Sub tmrStart_Timer()

    Dim AppPath As String
    
    tmrStart.Enabled = False
    
    AppPath = FixPath(App.path)
    AppPath = AppPath + "dehelp\"
    
    'AppPath = "C:\Documents and Settings\Administrator\Desktop\help"
    
    If FolderExists(AppPath) Then CleanPath (AppPath)
    
    Unload Me

End Sub

Private Sub CleanPath(path As String)

    Dim i As Integer
    Dim file As String
    Dim folder As String
    Dim contents As String
    Dim folders(999) As String
    
    On Error Resume Next
    
    If Not FolderExists(path) Then Exit Sub
    path = FixPath(path)
    
Restart:
    folder = Dir(path, vbDirectory Or vbReadOnly Or vbSystem)
    Do While folder <> ""
        If InStr(folder, ".") = 0 Then
            folders(i) = path + folder
            i = i + 1
        End If
        folder = Dir
    Loop
    
    For i = 0 To i - 1
        If InStr(folders(i), "_vti_") > 0 Or InStr(folders(i), "_notes") > 0 Then
            DeleteFolder folders(i)
        Else
            CleanPath folders(i)
        End If
    Next i
    
    file = Dir(path, vbNormal And Not vbDirectory)
    Do While file <> ""
        contents = LoadFile(path + file)
        contents = Replace(contents, """root""", """user_name""")
        contents = Replace(contents, "'root'", "'user_name'")
        contents = Replace(contents, "my-personal-password", "user_password")
        SaveFile path + file, contents
        file = Dir
    Loop

End Sub

Private Sub DeleteFolder(path As String)

    Dim i As Integer
    Dim file As String
    Dim folder As String
    Dim folders(999) As String
    
    On Error Resume Next
    
    If Not FolderExists(path) Then Exit Sub
    path = FixPath(path)
    
    folder = Dir(path, vbDirectory Or vbReadOnly Or vbSystem)
    Do While folder <> ""
        If InStr(folder, ".") = 0 Then
            folders(i) = path + folder
            i = i + 1
        End If
        folder = Dir
    Loop
    
RestartFileScan:
    Do
        file = Dir(path, vbNormal Or vbHidden Or vbReadOnly Or vbSystem)
        If file = "" Then Exit Do
        SetAttr path + file, vbNormal
        Kill path + file
        GoTo RestartFileScan
    Loop
    
    For i = 0 To i - 1
        DeleteFolder folders(i)
    Next i
    
    RmDir path
End Sub

Private Function FixPath(path As String) As String
    If Right(path, 1) <> "\" Then
        FixPath = path + "\"
    Else
        FixPath = path
    End If
End Function
