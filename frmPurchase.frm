VERSION 5.00
Object = "{2A2AD7CA-AC77-46F3-84DC-115021432312}#1.0#0"; "HREF.OCX"
Begin VB.Form frmPurchase 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Purchase DHTML Menu Builder"
   ClientHeight    =   5700
   ClientLeft      =   7215
   ClientTop       =   4770
   ClientWidth     =   6885
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
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
   ScaleHeight     =   5700
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frameReseller 
      Caption         =   "Select Reseller"
      Height          =   2100
      Left            =   45
      TabIndex        =   12
      Top             =   3015
      Width           =   6780
      Begin VB.PictureBox imgR 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   600
         Index           =   2
         Left            =   5670
         ScaleHeight     =   600
         ScaleWidth      =   540
         TabIndex        =   19
         Top             =   585
         Width           =   540
      End
      Begin VB.PictureBox imgR 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   600
         Index           =   1
         Left            =   3060
         ScaleHeight     =   600
         ScaleWidth      =   540
         TabIndex        =   18
         Top             =   405
         Width           =   540
      End
      Begin VB.PictureBox imgR 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   600
         Index           =   0
         Left            =   435
         ScaleHeight     =   600
         ScaleWidth      =   540
         TabIndex        =   17
         Top             =   435
         Width           =   540
      End
      Begin VB.OptionButton opReseller 
         Height          =   210
         Index           =   2
         Left            =   5790
         TabIndex        =   15
         Top             =   1440
         Width           =   210
      End
      Begin VB.OptionButton opReseller 
         Enabled         =   0   'False
         Height          =   210
         Index           =   1
         Left            =   3225
         TabIndex        =   14
         Top             =   1440
         Width           =   210
      End
      Begin VB.OptionButton opReseller 
         Enabled         =   0   'False
         Height          =   210
         Index           =   0
         Left            =   660
         TabIndex        =   13
         Top             =   1440
         Value           =   -1  'True
         Width           =   210
      End
      Begin href.uchref1 lnkResellerInfo 
         Height          =   315
         Left            =   60
         TabIndex        =   16
         Top             =   1740
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   556
         Caption         =   "Resellers Information..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Tahoma"
         FontSize        =   8.25
         ForeColor       =   16711680
         URL             =   "http://software.xfx.net/utilities/dmbuilder/purchase.php"
      End
      Begin VB.Shape shpSel 
         BorderColor     =   &H80000002&
         BorderWidth     =   2
         Height          =   1050
         Left            =   1680
         Top             =   555
         Width           =   1080
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Select the number of units/licenses"
      Height          =   1620
      Left            =   3750
      TabIndex        =   7
      Top             =   1290
      Width           =   3075
      Begin href.uchref1 lnkLicenseInfo 
         Height          =   315
         Left            =   60
         TabIndex        =   10
         ToolTipText     =   "Click here to display pricing information for license type"
         Top             =   1275
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         Caption         =   "Unit/License Information..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Tahoma"
         FontSize        =   8.25
         ForeColor       =   16711680
         URL             =   ""
      End
      Begin VB.OptionButton opLicense 
         Caption         =   "1 or more"
         Height          =   225
         Index           =   0
         Left            =   255
         TabIndex        =   9
         Top             =   300
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.OptionButton opLicense 
         Caption         =   "Site License"
         Height          =   225
         Index           =   1
         Left            =   255
         TabIndex        =   8
         Top             =   630
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Version"
      Height          =   1620
      Left            =   45
      TabIndex        =   3
      Top             =   1290
      Width           =   3630
      Begin VB.OptionButton opVersion 
         Caption         =   "Developers' Edition"
         Enabled         =   0   'False
         Height          =   225
         Index           =   2
         Left            =   255
         TabIndex        =   6
         Top             =   960
         Width           =   2295
      End
      Begin VB.OptionButton opVersion 
         Caption         =   "LITE Edition"
         Enabled         =   0   'False
         Height          =   225
         Index           =   1
         Left            =   255
         TabIndex        =   5
         Top             =   630
         Width           =   2295
      End
      Begin VB.OptionButton opVersion 
         Caption         =   "Standard Edition"
         Height          =   225
         Index           =   0
         Left            =   255
         TabIndex        =   4
         Top             =   300
         Value           =   -1  'True
         Width           =   2295
      End
      Begin href.uchref1 lnkVerInfo 
         Height          =   315
         Left            =   60
         TabIndex        =   11
         ToolTipText     =   "Click here to display a chart comparing each version"
         Top             =   1275
         Width           =   1710
         _ExtentX        =   3016
         _ExtentY        =   556
         Caption         =   "Version Information..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Tahoma"
         FontSize        =   8.25
         ForeColor       =   16711680
         URL             =   ""
      End
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   345
      Left            =   5565
      TabIndex        =   2
      Top             =   5295
      Width           =   1260
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Continue »"
      Default         =   -1  'True
      Height          =   345
      Left            =   4140
      TabIndex        =   1
      Top             =   5295
      Width           =   1260
   End
   Begin VB.PictureBox picLogo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1155
      Left            =   15
      ScaleHeight     =   1155
      ScaleWidth      =   6855
      TabIndex        =   0
      Top             =   15
      Width           =   6855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "For pricing information and to continue with the purchase process press the Continue button"
      Height          =   390
      Left            =   45
      TabIndex        =   20
      Top             =   5235
      Width           =   3720
   End
End
Attribute VB_Name = "frmPurchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()

    Unload Me

End Sub

Private Sub cmdOK_Click()

    Dim sUrl As String
    
    If ResellerID = "" Then
        If opReseller(0).Value Then
            sUrl = "http://www.reg.net/product.asp?id="
            If opVersion(0).Value Then
                If opLicense(0).Value Then sUrl = sUrl + "5499"
                If opLicense(1).Value Then sUrl = sUrl + "9767"
            End If
            If opVersion(1).Value Then sUrl = sUrl + "6267"
        End If
        If opReseller(1).Value Then
            sUrl = "http://www.regnow.com/softsell/nph-softsell.cgi?item=3791-"
            If opVersion(0).Value Then
                If opLicense(0).Value Then sUrl = sUrl + "1"
                If opLicense(1).Value Then sUrl = sUrl + "2"
            End If
            If opVersion(1).Value Then sUrl = sUrl + "3"
        End If
        If opReseller(2).Value Then
            If opVersion(0).Value Then sUrl = "https://xfx.net/utilities/dmbuilder/paypal.htm"
            If opVersion(1).Value Then sUrl = "https://xfx.net/utilities/dmbuilderlite/paypal.htm"
        End If
    Else
        If opVersion(0).Value Then
            If opLicense(0).Value Then sUrl = ResellerInfo(3)
            If opLicense(1).Value Then sUrl = ResellerInfo(5)
        End If
        
        If opVersion(1).Value Then sUrl = ResellerInfo(4)
    End If
    
    RunShellExecute "Open", sUrl, 0, 0, 0
    
    DoEvents
    
    Unload Me

End Sub

Private Sub Form_Load()

    CenterForm Me
    FixCtrls4Skin Me

    Set picLogo.Picture = LoadPicture(AppPath + "rsc\about.jpg")
    
    If ResellerID = "" Then
        Set imgR(0).Picture = LoadPicture(AppPath + "rinfo\regnet_logo.gif")
        Set imgR(1).Picture = LoadPicture(AppPath + "rinfo\regnow_logo.gif")
        Set imgR(2).Picture = LoadPicture(AppPath + "rinfo\paypal_logo.gif")
        
        imgR(0).BackColor = Me.BackColor
        imgR(1).BackColor = Me.BackColor
        imgR(2).BackColor = Me.BackColor
    Else
        Set imgR(0).Picture = LoadPicture(AppPath + "rinfo\" + ResellerID + "\logo.gif")
        
        imgR(0).BackColor = IIf(ResellerInfo(6) = -1, Me.BackColor, ResellerInfo(6))
        
        lnkResellerInfo.caption = "DHTML Menu Builder " + ResellerInfo(8)
        lnkResellerInfo.url = IIf(ResellerInfo(9) = "", ResellerInfo(1), ResellerInfo(9))
        
        opReseller(1).Visible = False
        opReseller(2).Visible = False
        imgR(1).Visible = False
        imgR(2).Visible = False
    End If
    CenterRImg 0
    CenterRImg 1
    CenterRImg 2
    
    opReseller_Click 0
    opReseller(2).Value = True

End Sub

Private Sub CenterRImg(idx As Integer)

    With opReseller(idx)
        imgR(idx).Left = .Left + (.Width - imgR(idx).Width) / 2
        imgR(idx).Top = (.Top + IIf(IsSkinned, 0, 15 * Screen.TwipsPerPixelY) - imgR(idx).Height) / 2
    End With

End Sub

Private Sub imgR_Click(Index As Integer)

    opReseller(Index).Value = True
    
    If opVersion(0).Value Then opVersion_Click 0
    If opVersion(1).Value Then opVersion_Click 1
    If opVersion(2).Value Then opVersion_Click 2

End Sub

Private Sub lnkLicenseInfo_Click()

    showHelp "https://xfx.net/utilities/dmbuilder/purchase.php?dataOnly=1"

End Sub

Private Sub lnkResellerInfo_Click()

    RunShellExecute "Open", lnkResellerInfo.url, 0, 0, 0

End Sub

Private Sub lnkVerInfo_Click()

    showHelp "https://xfx.net/utilities/dmbuilderlite/stdvsse.php?dataOnly=1"

End Sub

Private Sub opReseller_Click(Index As Integer)

    With shpSel
        .Left = imgR(Index).Left - 2 * Screen.TwipsPerPixelX
        If IsSkinned Then
            .Top = 1 * Screen.TwipsPerPixelY
        Else
            .Top = 20 * Screen.TwipsPerPixelY
        End If
        .Width = imgR(Index).Width + 5 * Screen.TwipsPerPixelX
        .Height = opReseller(0).Top + opReseller(0).Height + 5 * Screen.TwipsPerPixelY - .Top
    End With

End Sub

Private Sub opVersion_Click(Index As Integer)

    If Index > 0 Or (opReseller(2).Value = True) Then
        opLicense(0).Value = True
        opLicense(0).caption = "1"
    Else
        opLicense(0).caption = "1 or more"
    End If
    opLicense(1).Enabled = (Index = 0) And (opReseller(2).Value = False)
    
    If Index = 2 Then
        MsgBox "The Developers' Edition is not yet available for purchase." + vbCrLf + _
            "If you're interested in this version, purchase the Standard Edition and you will be able to register the Developers' Edition with your Standard Edition's license.", vbInformation + vbOKOnly, "Developers' Edition Information"
            opVersion(0).Value = True
    End If

End Sub
