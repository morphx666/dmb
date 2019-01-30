VERSION 5.00
Begin VB.Form frmSerialize 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "License"
   ClientHeight    =   2715
   ClientLeft      =   4560
   ClientTop       =   4815
   ClientWidth     =   5775
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
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtName 
      Height          =   315
      Left            =   2820
      TabIndex        =   0
      Top             =   840
      Width           =   2250
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4650
      TabIndex        =   7
      Top             =   2250
      Width           =   1050
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   2250
      Width           =   1050
   End
   Begin VB.TextBox txtLicense 
      Height          =   315
      Left            =   2820
      TabIndex        =   1
      Top             =   1560
      Width           =   2250
   End
   Begin VB.TextBox txtSerial 
      Height          =   315
      Left            =   2820
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   1200
      Width           =   2250
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Your Name"
      Height          =   210
      Left            =   1800
      TabIndex        =   10
      Top             =   885
      Width           =   915
   End
   Begin VB.Label lblWebPage 
      AutoSize        =   -1  'True
      Caption         =   "http://software.xfx.net"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   60
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   2490
      Width           =   1725
   End
   Begin VB.Label Label1 
      Caption         =   "For information on how to register this product, please go to"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   45
      TabIndex        =   8
      Top             =   2100
      Width           =   2550
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   195
      Picture         =   "frmSerialize.frx":0000
      Top             =   180
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "License Code"
      Height          =   210
      Left            =   1635
      TabIndex        =   5
      Top             =   1605
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Product Serial Number"
      Height          =   210
      Left            =   885
      TabIndex        =   4
      Top             =   1245
      Width           =   1830
   End
   Begin VB.Label lblMsg 
      Caption         =   "msg"
      Height          =   450
      Left            =   885
      TabIndex        =   2
      Top             =   195
      Width           =   4650
   End
End
Attribute VB_Name = "frmSerialize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdOK_Click()

    If txtName.Text = "" Then
        MsgBox "Your full name is requiered in order to register this product"
    Else
        SaveSetting App.EXEName, "Config", "Serial", txtLicense.Text
        SaveSetting App.EXEName, "Config", "User", txtName.Text
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()

    With Screen
        Left = .Width / 2 - Width / 2
        Top = .Height / 2 - Height / 2
    End With

    txtSerial.Text = EncryptHDSerial

End Sub

Private Function EncryptHDSerial() As String

    Dim volName As String
    Dim drvSerial As String
    Dim i As Integer

    rgbGetVolume "c:\", volName, drvSerial
    
    For i = Len(drvSerial) To 1 Step -1
        If Int(i / 2) = i / 2 Then
            EncryptHDSerial = EncryptHDSerial + Left(Asc(Mid(drvSerial, i, 1)), 1) + Right(Asc(Mid(drvSerial, i, 1)), 1)
        Else
            EncryptHDSerial = EncryptHDSerial + Right(Asc(Mid(drvSerial, i, 1)), 1) + Left(Asc(Mid(drvSerial, i, 1)), 1)
        End If
    Next i

End Function

Private Sub lblWebPage_Click()

    Shell "start http://software.xfx.net", vbHide

End Sub
