VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmFramesInfo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Retrieving Frames Information"
   ClientHeight    =   1560
   ClientLeft      =   6570
   ClientTop       =   7665
   ClientWidth     =   5580
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   11  'Hourglass
   Moveable        =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   5580
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrInit 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   300
      Top             =   915
   End
   Begin VB.Timer tmrFailed 
      Interval        =   30000
      Left            =   165
      Top             =   765
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4620
      MousePointer    =   1  'Arrow
      TabIndex        =   0
      Top             =   1110
      Width           =   900
   End
   Begin SHDocVwCtl.WebBrowser wbCtrl 
      Height          =   2010
      Left            =   5910
      TabIndex        =   2
      Top             =   60
      Width           =   2130
      ExtentX         =   3757
      ExtentY         =   3545
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin MSComctlLib.ProgressBar pbProgress 
      Height          =   300
      Left            =   945
      TabIndex        =   4
      Top             =   705
      Width           =   4560
      _ExtentX        =   8043
      _ExtentY        =   529
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "-------------"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   945
      TabIndex        =   3
      Top             =   390
      Width           =   4515
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   180
      Picture         =   "frmFramesInfo.frx":0000
      Top             =   255
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Loading Document"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   945
      TabIndex        =   1
      Top             =   135
      Width           =   1320
   End
End
Attribute VB_Name = "frmFramesInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

    FramesInfo.IsValid = False
    Unload Me

End Sub

Private Sub Form_Load()

    Width = 5700
    Height = 1905
    
    DoEvents

    CenterForm Me
    SetupCharset Me
    
    pbProgress.Max = 100
    pbProgress.value = 0
    
    lblInfo.Caption = EllipseText(lblInfo, FramesInfo.FileName, DT_WORD_ELLIPSIS)
    
    tmrInit.Enabled = True
    
End Sub

Private Sub tmrFailed_Timer()

    tmrFailed.Enabled = False

    If MsgBox("The process is taking more time than expected. Depending on the contents of your documents this process may require more time." + vbCrLf + _
              "Would you like to cancel it now?", vbInformation + vbYesNo, "DHTML Menu Builder") = vbYes Then
        cmdCancel_Click
    Else
        tmrFailed.Enabled = True
    End If

End Sub

Private Sub tmrInit_Timer()

    On Error GoTo Abort

    Dim IsValid As Boolean
    tmrInit.Enabled = False
    
    If Mid(FramesInfo.FileName, 2, 1) = ":" Or Left(FramesInfo.FileName, 2) = "\\" Then
        If Dir(FramesInfo.FileName) <> "" And FramesInfo.FileName <> "" Then
            If Not FileIsHTML(FramesInfo.FileName) Then
                MsgBox GetLocalizedStr(676), vbInformation + vbOKOnly, GetLocalizedStr(586)
                IsValid = False
            Else
                IsValid = True
            End If
        End If
    Else
        IsValid = True
    End If
    
    If IsValid Then
        With wbCtrl
            .Silent = True
            .Navigate FramesInfo.FileName
        End With
        Exit Sub
    End If
    
Abort:
    
    cmdCancel_Click

End Sub

Private Sub wbCtrl_DocumentComplete(ByVal pDisp As Object, URL As Variant)

    Dim doc As IHTMLDocument2
    
    If LCase(URL) <> LCase(FramesInfo.FileName) Or URL = "" Then Exit Sub
    
    Do While wbCtrl.Document Is Nothing
        DoEvents
    Loop

    Set doc = wbCtrl.Document
    
    ReDim FramesInfo.Frames(0)
    GetFrames doc.Frames, "top."
    
    FramesInfo.IsValid = (UBound(FramesInfo.Frames) > 0)
    
    Unload Me
    
End Sub

Private Sub GetFrames(fCol As Object, fPath As String)

    Dim i As Integer
    Dim nPath As String
    
    On Error GoTo chkErr
    
    For i = 0 To fCol.Length - 1
        nPath = fPath + fCol.item(i).Name
        
        ReDim Preserve FramesInfo.Frames(UBound(FramesInfo.Frames) + 1)
        FramesInfo.Frames(UBound(FramesInfo.Frames)) = nPath

        If fCol.item(i).Length > 0 Then
            GetFrames fCol.item(i), nPath + "."
        End If
ContinueAfterError:
    Next i

    Exit Sub
    
chkErr:
    MsgBox "Unable to get frames information" + vbCrLf + "Error " & Err.Number & ": " + Err.Description + vbCrLf + vbCrLf + "DHTML Menu Builder will try to retrieve the rest of the frames information", vbCritical + vbOKOnly, GetLocalizedStr(664)
    Resume ContinueAfterError

End Sub

Private Sub wbCtrl_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)

    On Error Resume Next

    pbProgress.Max = ProgressMax
    pbProgress.value = Progress

End Sub
