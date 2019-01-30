VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Preset Installer"
   ClientHeight    =   2175
   ClientLeft      =   3435
   ClientTop       =   4395
   ClientWidth     =   6585
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   405
      Left            =   5355
      TabIndex        =   1
      Top             =   1695
      Width           =   1155
   End
   Begin VB.Timer tmrDoCopy 
      Interval        =   50
      Left            =   390
      Top             =   1620
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   225
      Picture         =   "frmMain.frx":2CFA
      Top             =   585
      Width           =   480
   End
   Begin VB.Label lblMessage 
      BackStyle       =   0  'Transparent
      Caption         =   "Results..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1410
      Left            =   1080
      TabIndex        =   0
      Top             =   135
      Width           =   5430
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    Move Screen.Width / 2 - Width / 2, Screen.Height / 2 - Height / 2

End Sub

Private Sub tmrDoCopy_Timer()

    tmrDoCopy.Enabled = False
    
    If Command = "" Then
        Unload Me
    Else
        DoCopy
    End If

End Sub

Private Function GetExtension(FileName As String) As String

    If InStr(FileName, ".") Then
        GetExtension = LCase(Mid(FileName, InStrRev(FileName, ".") + 1))
    End If

End Function

Private Sub DoCopy()
    
    Dim Failed As Boolean
    Dim SrcFile As String
    Dim DstPath As String
    Dim PresetFile As String
    Dim Ans As Integer
    
    On Error GoTo ShowError
    
    SrcFile = Command
    If Right(SrcFile, 1) = Chr(34) Then SrcFile = Mid(SrcFile, 2, Len(SrcFile) - 2)
    DstPath = QueryValue(HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\DMB\RegInfo", "InstallPath")
    If Right(DstPath, 1) <> "\" Then DstPath = DstPath + "\"

    If DstPath = "\" Or Dir(DstPath) = "" Then
        Failed = True
        lblMessage.Caption = "Error Installing " + GetFileName(SrcFile) + vbCrLf + vbCrLf + "The DHTML Menu Builder application could not be found and its required to use this Preset"
    Else
        If GetExtension(SrcFile) <> "dpp" Then
            Failed = True
            lblMessage.Caption = "Error Installing " + GetFileName(SrcFile) + vbCrLf + vbCrLf + "The selected file is not a valid DHTML Menu Builder Preset"
        End If
    End If
    
    If Not Failed Then
        PresetFile = GetFileName(Short2Long(SrcFile))
        If Dir(DstPath + "Presets\" + PresetFile) <> "" Then
            Ans = MsgBox("The Preset " + PresetFile + " already exists, do you want to overwrite it?", vbQuestion + vbYesNo, "Warning")
        Else
            Ans = vbYes
        End If
        If Ans = vbYes Then
            If Not FileExists(SrcFile) Then
                MsgBox "The Preset " + SrcFile + " could not be found"
            Else
                FileCopy SrcFile, DstPath + "Presets\" + PresetFile
          
                lblMessage.Caption = PresetFile + " has been successfully installed." + vbCrLf + vbCrLf + "To use this Preset in DHTML Menu Builder, click File->New->From Preset or click Tools->Apply Style from Preset"
                cmdClose.Caption = "OK"
                On Error Resume Next
                'Kill SrcFile
            End If
        Else
            lblMessage.Caption = vbCrLf + vbCrLf + "Installation aborted..."
            cmdClose.Caption = "&Close"
        End If
    Else
        cmdClose.Caption = "&Close"
    End If
    
    Exit Sub
    
ShowError:
    lblMessage.Caption = "An internal error " & Err.Number & " has occured (" & Erl & "): " + vbCrLf + _
      Err.Description & vbCrLf & _
      "Source: " & SrcFile & vbCrLf & _
      "Destination: " & DstPath + "Presets\" + GetFileName(DecodeUrl(Short2Long(PresetFile)))
  
    cmdClose.Caption = "&Close"

End Sub

Private Function chExt(ByVal FileName As String, NewExt As String) As String

    Mid(FileName, InStrRev(FileName, ".") + 1) = NewExt
    
    chExt = FileName

End Function

Private Function GetFileName(ByVal FileName As String) As String

    Dim p As Integer

    FileName = Replace(FileName, "/", "\")
    p = InStrRev(FileName, "\")
    If p > 0 Then
        GetFileName = Mid$(FileName, p + 1)
    Else
        GetFileName = ""
    End If

End Function
