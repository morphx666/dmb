VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DHTML Menu Builder RegIt"
   ClientHeight    =   1125
   ClientLeft      =   5595
   ClientTop       =   6000
   ClientWidth     =   4830
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
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
   ScaleHeight     =   1125
   ScaleWidth      =   4830
   Begin VB.Timer tmrClose 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   3015
      Top             =   585
   End
   Begin VB.Timer tmrStart 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   855
      Top             =   525
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Please wait while Registering DHTML Menu Builder"
      Height          =   405
      Left            =   1020
      TabIndex        =   0
      Top             =   300
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   180
      Picture         =   "frmMain.frx":0CCA
      Top             =   255
      Width           =   480
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Form_Load()

    With Screen
        Move (.Width - Width) \ 2, (.Height - Height) \ 2
    End With
    
    tmrStart.Enabled = True

End Sub

Private Sub RegIt()

    Dim USER As String
    Dim COMPANY As String
    Dim USERSN As String
    
    Dim p As String
    Dim c As String
        
    If GetSetting("DMB", "RegInfo", "PreRegVer", 1) / GetSetting("DMB", "RegInfo", "CacheSig02", 1) = 7 Then
        USER = GetSetting("DMB", "RegInfo", "User", "DEMO")
        COMPANY = GetSetting("DMB", "RegInfo", "Company", "DEMO")
        USERSN = GetSetting("DMB", "RegInfo", "SerialNumber", "")
    
        p = USER + "|" + COMPANY + "|" + USERSN + "|" + GetHDSerial + "|" + "STD|"
        c = Str2HEX(p)
        If RunShellExecute("Open", "register.exe", c, Long2Short(App.Path), 1) <= 32 Then
            MsgBox "Error 1001: An unknown error has occurred while trying to register DHTML Menu Builder.", vbCritical + vbOKOnly, "Error Launching Register.exe"
        Else
            c = Str2HEX(Deflate(c))
            SaveSetting "DMB", "RegInfo", "CacheData", c
            SaveSetting "DMB", "RegInfo", "CacheSig01", Month(Now)
            SaveSetting "DMB", "RegInfo", "CacheSig03", "1"
    
            MsgBox "DHTML Menu Builder has been successfully registered. You may start the application by clicking on the shortcut icon on your Desktop.", vbInformation + vbOKOnly, "Registration"
        End If
    Else
        MsgBox "Error 1002: An unknown error has occurred while trying to register DHTML Menu Builder.", vbCritical + vbOKOnly, "Error Launching Register.exe"
    End If
    tmrClose.Enabled = True

End Sub

Public Function RunShellExecute(sTopic As String, sFile As Variant, sParams As Variant, sDirectory As Variant, nShowCmd As Long) As Long

   Dim hWndDesk As Long
   Dim success As Long
  
  'the desktop will be the
  'default for error messages
   hWndDesk = GetDesktopWindow()
  
  'execute the passed operation
   success = ShellExecute(hWndDesk, sTopic, sFile, sParams, sDirectory, nShowCmd)

  'This is optional. Uncomment the three lines
  'below to have the "Open With.." dialog appear
  'when the ShellExecute API call fails
  'If success = SE_ERR_NOASSOC Then
  '   Call Shell("rundll32.exe shell32.dll,OpenAs_RunDLL " & sFile, vbNormalFocus)
  'End If
  
  RunShellExecute = success
   
End Function

Private Sub tmrClose_Timer()

    tmrClose.Enabled = False
    
    Unload Me

End Sub

Private Sub tmrStart_Timer()

    tmrStart.Enabled = True
    RegIt

End Sub
