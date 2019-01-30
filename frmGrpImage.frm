VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E1C6DB9D-BD4A-4A61-A759-0CED75D034BF}#43.0#0"; "SmartButton.ocx"
Object = "{9A2947B4-87EE-40EE-A3EF-32BDC32D5726}#1.0#0"; "xfxline3d.ocx"
Begin VB.Form frmGrpImage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Image"
   ClientHeight    =   6540
   ClientLeft      =   8145
   ClientTop       =   5700
   ClientWidth     =   7500
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGrpImage.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   Begin SmartButtonProject.SmartButton sbApplyOptions 
      Height          =   345
      Left            =   135
      TabIndex        =   61
      Top             =   120
      Width           =   7230
      _ExtentX        =   12753
      _ExtentY        =   609
      Caption         =   "Options"
      Picture         =   "frmGrpImage.frx":014A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionLayout   =   3
      PictureLayout   =   3
   End
   Begin VB.PictureBox picSize 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   1305
      ScaleHeight     =   675
      ScaleWidth      =   735
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   5850
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame frameTBItem 
      BorderStyle     =   0  'None
      Height          =   3450
      Left            =   90
      TabIndex        =   21
      Top             =   1305
      Width           =   7125
      Begin VB.Frame frameLeft 
         Caption         =   "Left Image"
         Height          =   3330
         Left            =   15
         TabIndex        =   22
         Top             =   90
         Width           =   2235
         Begin VB.TextBox txtLeftM 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1575
            TabIndex        =   65
            Text            =   "000"
            Top             =   2880
            Width           =   420
         End
         Begin VB.Frame frameNL 
            Caption         =   "Normal"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   195
            TabIndex        =   23
            Top             =   315
            Width           =   1800
            Begin VB.PictureBox picLeftNormal 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               ForeColor       =   &H80000008&
               Height          =   480
               Left            =   105
               ScaleHeight     =   450
               ScaleWidth      =   450
               TabIndex        =   24
               TabStop         =   0   'False
               Top             =   375
               Width           =   480
            End
            Begin SmartButtonProject.SmartButton cmdChange 
               Height          =   240
               Index           =   0
               Left            =   630
               TabIndex        =   25
               Top             =   375
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   423
               Caption         =   "Change"
               Picture         =   "frmGrpImage.frx":02A4
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CaptionLayout   =   4
               PictureLayout   =   3
            End
            Begin SmartButtonProject.SmartButton cmdRemoveLeft 
               Height          =   240
               Left            =   630
               TabIndex        =   26
               Top             =   615
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   423
               Caption         =   "Remove"
               Picture         =   "frmGrpImage.frx":063E
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CaptionLayout   =   4
               PictureLayout   =   3
            End
         End
         Begin VB.Frame frameOL 
            Caption         =   "Mouse Over"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   195
            TabIndex        =   27
            Top             =   1470
            Width           =   1800
            Begin VB.PictureBox picLeftOver 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               ForeColor       =   &H80000008&
               Height          =   480
               Left            =   105
               ScaleHeight     =   450
               ScaleWidth      =   450
               TabIndex        =   28
               TabStop         =   0   'False
               Top             =   375
               Width           =   480
            End
            Begin SmartButtonProject.SmartButton cmdChange 
               Height          =   240
               Index           =   1
               Left            =   630
               TabIndex        =   29
               Top             =   375
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   423
               Caption         =   "Change"
               Picture         =   "frmGrpImage.frx":09D8
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CaptionLayout   =   4
               PictureLayout   =   3
            End
            Begin SmartButtonProject.SmartButton cmdSameLeft 
               Height          =   240
               Left            =   630
               TabIndex        =   30
               Top             =   615
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   423
               Caption         =   "Same"
               Picture         =   "frmGrpImage.frx":0D72
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CaptionLayout   =   4
               PictureLayout   =   3
            End
         End
         Begin VB.TextBox txtLeftH 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   810
            TabIndex        =   34
            Text            =   "000"
            Top             =   2880
            WhatsThisHelpID =   20340
            Width           =   420
         End
         Begin VB.TextBox txtLeftW 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   195
            TabIndex        =   32
            Text            =   "000"
            Top             =   2880
            Width           =   420
         End
         Begin VB.Label lblLeftM 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Margin"
            Height          =   210
            Left            =   1470
            TabIndex        =   66
            Top             =   2655
            Width           =   525
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "x"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   660
            TabIndex        =   33
            Top             =   2910
            Width           =   105
         End
         Begin VB.Label lblLeftS 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Size"
            Height          =   210
            Left            =   195
            TabIndex        =   31
            Top             =   2655
            Width           =   315
         End
      End
      Begin VB.Frame frameRight 
         Caption         =   "Right Image"
         Height          =   3330
         Left            =   4785
         TabIndex        =   44
         Top             =   90
         Width           =   2235
         Begin VB.TextBox txtRightM 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1575
            TabIndex        =   67
            Text            =   "000"
            Top             =   2880
            Width           =   420
         End
         Begin VB.Frame frameOR 
            Caption         =   "Mouse Over"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   195
            TabIndex        =   49
            Top             =   1470
            Width           =   1800
            Begin VB.PictureBox picRightOver 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               ForeColor       =   &H80000008&
               Height          =   480
               Left            =   105
               ScaleHeight     =   450
               ScaleWidth      =   450
               TabIndex        =   50
               TabStop         =   0   'False
               Top             =   375
               Width           =   480
            End
            Begin SmartButtonProject.SmartButton cmdChange 
               Height          =   240
               Index           =   3
               Left            =   630
               TabIndex        =   51
               Top             =   375
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   423
               Caption         =   "Change"
               Picture         =   "frmGrpImage.frx":110C
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CaptionLayout   =   4
               PictureLayout   =   3
            End
            Begin SmartButtonProject.SmartButton cmdSameRight 
               Height          =   240
               Left            =   630
               TabIndex        =   52
               Top             =   615
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   423
               Caption         =   "Same"
               Picture         =   "frmGrpImage.frx":14A6
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CaptionLayout   =   4
               PictureLayout   =   3
            End
         End
         Begin VB.TextBox txtRightW 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   195
            TabIndex        =   54
            Text            =   "000"
            Top             =   2880
            Width           =   420
         End
         Begin VB.TextBox txtRightH 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   810
            TabIndex        =   56
            Text            =   "000"
            Top             =   2880
            WhatsThisHelpID =   20340
            Width           =   420
         End
         Begin VB.Frame frameNR 
            Caption         =   "Normal"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   195
            TabIndex        =   45
            Top             =   315
            Width           =   1800
            Begin VB.PictureBox picRightNormal 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               ForeColor       =   &H80000008&
               Height          =   480
               Left            =   105
               ScaleHeight     =   450
               ScaleWidth      =   450
               TabIndex        =   46
               TabStop         =   0   'False
               Top             =   375
               Width           =   480
            End
            Begin SmartButtonProject.SmartButton cmdChange 
               Height          =   240
               Index           =   2
               Left            =   630
               TabIndex        =   47
               Top             =   375
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   423
               Caption         =   "Change"
               Picture         =   "frmGrpImage.frx":1840
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CaptionLayout   =   4
               PictureLayout   =   3
            End
            Begin SmartButtonProject.SmartButton cmdRemoveRight 
               Height          =   240
               Left            =   630
               TabIndex        =   48
               Top             =   615
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   423
               Caption         =   "Remove"
               Picture         =   "frmGrpImage.frx":1BDA
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CaptionLayout   =   4
               PictureLayout   =   3
            End
         End
         Begin VB.Label lblRightM 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Margin"
            Height          =   210
            Left            =   1470
            TabIndex        =   68
            Top             =   2655
            Width           =   525
         End
         Begin VB.Label lblRightS 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Size"
            Height          =   210
            Left            =   195
            TabIndex        =   53
            Top             =   2655
            Width           =   315
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "x"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   660
            TabIndex        =   55
            Top             =   2910
            Width           =   105
         End
      End
      Begin VB.Frame frameBack 
         Caption         =   "Background Image"
         Height          =   3330
         Left            =   2400
         TabIndex        =   35
         Top             =   90
         Width           =   2235
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   465
            Left            =   195
            ScaleHeight     =   465
            ScaleWidth      =   1800
            TabIndex        =   62
            Top             =   2700
            Width           =   1800
            Begin VB.CheckBox chkAllowCrop 
               Caption         =   "Allow Cropping"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   0
               TabIndex        =   64
               Top             =   255
               Width           =   1710
            End
            Begin VB.CheckBox chkTile 
               Caption         =   "Tile"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   0
               TabIndex        =   63
               Top             =   0
               Width           =   1710
            End
         End
         Begin VB.Frame frameNB 
            Caption         =   "Normal"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   195
            TabIndex        =   36
            Top             =   315
            Width           =   1800
            Begin VB.PictureBox picBackNormal 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               ForeColor       =   &H80000008&
               Height          =   480
               Left            =   105
               ScaleHeight     =   450
               ScaleWidth      =   450
               TabIndex        =   37
               TabStop         =   0   'False
               Top             =   375
               Width           =   480
            End
            Begin SmartButtonProject.SmartButton cmdChange 
               Height          =   240
               Index           =   5
               Left            =   630
               TabIndex        =   38
               Top             =   375
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   423
               Caption         =   "Change"
               Picture         =   "frmGrpImage.frx":1F74
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CaptionLayout   =   4
               PictureLayout   =   3
            End
            Begin SmartButtonProject.SmartButton cmdRemoveBack 
               Height          =   240
               Left            =   630
               TabIndex        =   39
               Top             =   615
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   423
               Caption         =   "Remove"
               Picture         =   "frmGrpImage.frx":230E
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CaptionLayout   =   4
               PictureLayout   =   3
            End
         End
         Begin VB.Frame frameOB 
            Caption         =   "Mouse Over"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   195
            TabIndex        =   40
            Top             =   1470
            Width           =   1800
            Begin VB.PictureBox picBackOver 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               ForeColor       =   &H80000008&
               Height          =   480
               Left            =   105
               ScaleHeight     =   450
               ScaleWidth      =   450
               TabIndex        =   41
               TabStop         =   0   'False
               Top             =   375
               Width           =   480
            End
            Begin SmartButtonProject.SmartButton cmdChange 
               Height          =   240
               Index           =   6
               Left            =   630
               TabIndex        =   42
               Top             =   375
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   423
               Caption         =   "Change"
               Picture         =   "frmGrpImage.frx":26A8
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CaptionLayout   =   4
               PictureLayout   =   3
            End
            Begin SmartButtonProject.SmartButton cmdSameBack 
               Height          =   240
               Left            =   630
               TabIndex        =   43
               Top             =   615
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   423
               Caption         =   "Same"
               Picture         =   "frmGrpImage.frx":2A42
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CaptionLayout   =   4
               PictureLayout   =   3
            End
         End
      End
   End
   Begin VB.Frame frameGroupProp 
      BorderStyle     =   0  'None
      Caption         =   "Group Properties"
      Height          =   3450
      Left            =   225
      TabIndex        =   0
      Top             =   930
      Width           =   7125
      Begin VB.PictureBox picCorner 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   3015
         ScaleHeight     =   450
         ScaleWidth      =   450
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1200
         Width           =   480
      End
      Begin SmartButtonProject.SmartButton sbCImage 
         Height          =   180
         Index           =   0
         Left            =   1500
         TabIndex        =   2
         Top             =   660
         Width           =   180
         _ExtentX        =   318
         _ExtentY        =   318
         BackColor       =   8421504
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.PictureBox picBackImage 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   2805
         ScaleHeight     =   450
         ScaleWidth      =   450
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   2760
         Width           =   480
      End
      Begin SmartButtonProject.SmartButton cmdChange 
         Height          =   240
         Index           =   4
         Left            =   1500
         TabIndex        =   16
         Top             =   2760
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   423
         Caption         =   "Change"
         Picture         =   "frmGrpImage.frx":2DDC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionLayout   =   4
         PictureLayout   =   3
      End
      Begin SmartButtonProject.SmartButton cmdRemove 
         Height          =   240
         Left            =   1500
         TabIndex        =   19
         Top             =   3000
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   423
         Caption         =   "Remove"
         Picture         =   "frmGrpImage.frx":3176
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionLayout   =   4
         PictureLayout   =   3
      End
      Begin xfxLine3D.ucLine3D uc3DLine1 
         Height          =   30
         Left            =   90
         TabIndex        =   15
         Top             =   2280
         Width           =   6780
         _ExtentX        =   11959
         _ExtentY        =   53
      End
      Begin SmartButtonProject.SmartButton sbCImage 
         Height          =   180
         Index           =   1
         Left            =   1680
         TabIndex        =   3
         Top             =   660
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   318
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SmartButtonProject.SmartButton sbCImage 
         Height          =   180
         Index           =   2
         Left            =   2220
         TabIndex        =   4
         Top             =   660
         Width           =   180
         _ExtentX        =   318
         _ExtentY        =   318
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SmartButtonProject.SmartButton sbCImage 
         Height          =   540
         Index           =   3
         Left            =   1500
         TabIndex        =   7
         Top             =   840
         Width           =   180
         _ExtentX        =   318
         _ExtentY        =   953
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SmartButtonProject.SmartButton sbCImage 
         Height          =   540
         Index           =   4
         Left            =   2220
         TabIndex        =   8
         Top             =   840
         Width           =   180
         _ExtentX        =   318
         _ExtentY        =   953
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SmartButtonProject.SmartButton sbCImage 
         Height          =   180
         Index           =   5
         Left            =   1500
         TabIndex        =   10
         Top             =   1380
         Width           =   180
         _ExtentX        =   318
         _ExtentY        =   318
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SmartButtonProject.SmartButton sbCImage 
         Height          =   180
         Index           =   6
         Left            =   1680
         TabIndex        =   11
         Top             =   1380
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   318
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SmartButtonProject.SmartButton sbCImage 
         Height          =   180
         Index           =   7
         Left            =   2220
         TabIndex        =   12
         Top             =   1380
         Width           =   180
         _ExtentX        =   318
         _ExtentY        =   318
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SmartButtonProject.SmartButton sbChangeCImage 
         Height          =   240
         Left            =   2655
         TabIndex        =   5
         Top             =   660
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   423
         Caption         =   "Change"
         Picture         =   "frmGrpImage.frx":3510
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionLayout   =   4
         PictureLayout   =   3
      End
      Begin SmartButtonProject.SmartButton sbRemoveCImage 
         Height          =   240
         Left            =   2655
         TabIndex        =   9
         Top             =   900
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   423
         Caption         =   "Remove"
         Picture         =   "frmGrpImage.frx":38AA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionLayout   =   4
         PictureLayout   =   3
      End
      Begin VB.Label lblCornerImages 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Corner Images"
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
         Left            =   195
         TabIndex        =   6
         Top             =   1080
         Width           =   1065
      End
      Begin VB.Label lblSelCImage 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Top Left"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   1500
         TabIndex        =   1
         Top             =   435
         Width           =   600
      End
      Begin VB.Label lblImages 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Images"
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
         Left            =   1695
         TabIndex        =   14
         Top             =   1635
         Width           =   525
      End
      Begin VB.Label lblBackImage 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Back Image"
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
         Left            =   435
         TabIndex        =   17
         Top             =   2895
         Width           =   825
      End
   End
   Begin VB.Frame frmLiveSample 
      Caption         =   "Live Sample"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1335
      Left            =   120
      TabIndex        =   20
      Top             =   4575
      Width           =   7260
   End
   Begin MSComctlLib.TabStrip tsOptions 
      Height          =   3915
      Left            =   120
      TabIndex        =   60
      Top             =   555
      Width           =   7260
      _ExtentX        =   12806
      _ExtentY        =   6906
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Group Properties"
            Key             =   "tsGroupProperties"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Toolbar Item"
            Key             =   "tsToolbarItem"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
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
      Left            =   5445
      TabIndex        =   58
      Top             =   6045
      Width           =   900
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
      Left            =   6480
      TabIndex        =   59
      Top             =   6045
      Width           =   900
   End
End
Attribute VB_Name = "frmGrpImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim BackGrp As MenuGrp
Dim SelId As Integer
Dim SelCIndex As Integer
Dim IsUpdating As Boolean

Private WithEvents msgSubClass As xfxSC
Attribute msgSubClass.VB_VarHelpID = -1

Private Sub chkAllowCrop_Click()

    On Error Resume Next
    If MenuGrps(SelId).tbiBackImage.NormalImage <> "" Then
        picSize.Picture = LoadPictureRes(MenuGrps(SelId).tbiBackImage.NormalImage)
        MenuGrps(SelId).tbiBackImage.w = picSize.Width / Screen.TwipsPerPixelX
        MenuGrps(SelId).tbiBackImage.h = picSize.Height / Screen.TwipsPerPixelY
    End If
    
    MenuGrps(SelId).tbiBackImage.AllowCrop = (chkAllowCrop.Value = vbChecked)
    UpdateSample

End Sub

Private Sub chkTile_Click()

    On Error Resume Next
    MenuGrps(SelId).tbiBackImage.Tile = (chkTile.Value = vbChecked)
    UpdateSample

End Sub

Private Sub cmdCancel_Click()

    MenuGrps(SelId) = BackGrp
    Unload Me

End Sub

Private Sub cmdChange_Click(Index As Integer)

    Select Case Index
        Case 0
            SelImage.FileName = MenuGrps(SelId).tbiLeftImage.NormalImage
        Case 1
            SelImage.FileName = MenuGrps(SelId).tbiLeftImage.HoverImage
        Case 2
            SelImage.FileName = MenuGrps(SelId).tbiRightImage.NormalImage
        Case 3
            SelImage.FileName = MenuGrps(SelId).tbiRightImage.HoverImage
        Case 4
            SelImage.FileName = MenuGrps(SelId).Image
            SelImage.SupportsFlash = True
        Case 5
            SelImage.FileName = MenuGrps(SelId).tbiBackImage.NormalImage
            SelImage.SupportsFlash = True
        Case 6
            SelImage.FileName = MenuGrps(SelId).tbiBackImage.HoverImage
            SelImage.SupportsFlash = True
    End Select
    frmRscImages.Show vbModal
    
    With SelImage
        If .IsValid Then
            Set picSize.Picture = LoadPictureRes(.FileName)
            DoEvents
            Select Case Index
                Case 0
                    txtLeftW.Text = picSize.Width / Screen.TwipsPerPixelX
                    txtLeftH.Text = picSize.Height / Screen.TwipsPerPixelY
                    With MenuGrps(SelId).tbiLeftImage
                        .NormalImage = SelImage.FileName
                        .w = picSize.Width / Screen.TwipsPerPixelX
                        .h = picSize.Height / Screen.TwipsPerPixelY
                    End With
                Case 1
                    MenuGrps(SelId).tbiLeftImage.HoverImage = .FileName
                Case 2
                    txtRightW.Text = picSize.Width / Screen.TwipsPerPixelX
                    txtRightH.Text = picSize.Height / Screen.TwipsPerPixelY
                    With MenuGrps(SelId).tbiRightImage
                        .NormalImage = SelImage.FileName
                        .w = picSize.Width / Screen.TwipsPerPixelX
                        .h = picSize.Height / Screen.TwipsPerPixelY
                    End With
                Case 3
                    MenuGrps(SelId).tbiRightImage.HoverImage = .FileName
                Case 4
                    MenuGrps(SelId).Image = .FileName
                Case 5
                    With MenuGrps(SelId).tbiBackImage
                        .NormalImage = SelImage.FileName
                        .w = picSize.Width / Screen.TwipsPerPixelX
                        .h = picSize.Height / Screen.TwipsPerPixelY
                    End With
                Case 6
                    MenuGrps(SelId).tbiBackImage.HoverImage = .FileName
            End Select
        End If
    End With
    
    Me.SetFocus
    UpdateSample

End Sub

Private Sub cmdOK_Click()

    ApplyStyleOptions
    frmMain.SaveState "Change " + MenuGrps(GetID).Name + " Image"
    
    Unload Me

End Sub

Private Sub ApplyStyleOptions()

    Dim i As Integer
    Dim c As Integer
    Dim t As Integer
    Dim sId As Integer
    
    sId = GetID
    
    For c = 0 To frmMain.mnuStyleOptionsOP.Count - 1
        If frmMain.mnuStyleOptionsOP(c).Checked Then
            t = Val(frmMain.mnuStyleOptionsOP(c).tag)
            Select Case c
                Case 0: ' do nothing
                Case 2:
                    For i = 1 To UBound(MenuGrps)
                        If BelongsToToolbar(i, True) = t Then CopyStyle sId, i
                    Next i
                Case 3:
                    For i = 1 To UBound(MenuGrps)
                        CopyStyle sId, i
                    Next i
            End Select
            Exit Sub
        End If
    Next c
    
    With dmbClipboard
        For i = 1 To UBound(.CustomSel)
            CopyStyle sId, GetIDByName(.CustomSel(i))
        Next i
    End With

End Sub

Private Sub CopyStyle(sId As Integer, tID As Integer)

    With MenuGrps(tID)
        .BackImage = MenuGrps(sId).BackImage
        .CornersImages = MenuGrps(sId).CornersImages
        .Image = MenuGrps(sId).Image
        .tbiBackImage = MenuGrps(sId).tbiBackImage
        .tbiLeftImage = MenuGrps(sId).tbiLeftImage
        .tbiRightImage = MenuGrps(sId).tbiRightImage
    End With

End Sub

Private Sub cmdRemove_Click()

    MenuGrps(SelId).Image = vbNullString
    
    UpdateSample

End Sub

Private Sub cmdRemoveBack_Click()

    With MenuGrps(SelId).tbiBackImage
        .NormalImage = vbNullString
        .HoverImage = vbNullString
    End With
    
    Me.SetFocus
    
    UpdateSample

End Sub

Private Sub cmdRemoveLeft_Click()

    With MenuGrps(SelId).tbiLeftImage
        .NormalImage = vbNullString
        .HoverImage = vbNullString
        .w = 0
        .h = 0
    End With
    
    Me.SetFocus
    
    UpdateSample

End Sub

Private Sub cmdSameBack_Click()

    With MenuGrps(SelId).tbiBackImage
        .HoverImage = .NormalImage
    End With
    
    Me.SetFocus
    
    UpdateSample

End Sub

Private Sub cmdSameLeft_Click()

    With MenuGrps(SelId).tbiLeftImage
        .HoverImage = .NormalImage
    End With
    
    Me.SetFocus
    
    UpdateSample

End Sub

Private Sub cmdRemoveRight_Click()

    With MenuGrps(SelId).tbiRightImage
        .NormalImage = vbNullString
        .HoverImage = vbNullString
        .w = 0
        .h = 0
    End With
    
    Me.SetFocus
    
    UpdateSample

End Sub

Private Sub cmdSameRight_Click()

    With MenuGrps(SelId).tbiRightImage
        .HoverImage = .NormalImage
    End With
    
    Me.SetFocus
    
    UpdateSample

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case tsOptions.SelectedItem.key
        Case "tsGroupProperties"
            If KeyCode = vbKeyF1 Then showHelp "dialogs/group_image.htm"
        Case "tsToolbarItem"
            If KeyCode = vbKeyF1 Then showHelp "dialogs/group_image_tbi.htm"
    End Select

End Sub

Private Sub msgSubClass_NewMessage(ByVal hwnd As Long, uMsg As Long, wParam As Long, lParam As Long, Cancel As Boolean, UseRetVal As Boolean, RetVal As Long)

    Select Case hwnd
        Case Me.hwnd
            Select Case uMsg
                Case WM_CTLCOLORSTATIC, WM_CTLCOLOREDIT, WM_PAINT
                    DrawColorBoxes Me
            End Select
    End Select
    
End Sub

Private Sub Form_Load()
    
    Dim i As Integer
    Dim DoTB As Boolean
    
    CenterForm Me
    LocalizeUI
    
    Set msgSubClass = New xfxSC
    msgSubClass.SubClassHwnd Me.hwnd, True
    
    SelId = GetID
    BackGrp = MenuGrps(SelId)
    
    caption = NiceGrpCaption(SelId) + " - " + GetLocalizedStr(214)
    
    DoTB = CreateToolbar
    frameTBItem.Enabled = DoTB
    frameNL.Enabled = DoTB
    frameNR.Enabled = DoTB
    frameOL.Enabled = DoTB
    frameOR.Enabled = DoTB
    For i = 0 To cmdChange.UBound
        cmdChange(i).Enabled = DoTB
    Next i
    cmdChange(4).Enabled = True
    cmdRemoveLeft.Enabled = DoTB
    cmdRemoveBack.Enabled = DoTB
    cmdRemoveRight.Enabled = DoTB
    cmdSameLeft.Enabled = DoTB
    cmdSameBack.Enabled = DoTB
    cmdSameRight.Enabled = DoTB
    txtLeftH.Enabled = DoTB
    txtLeftW.Enabled = DoTB
    txtRightH.Enabled = DoTB
    txtRightW.Enabled = DoTB
    lblLeftS.Enabled = DoTB
    lblRightS.Enabled = DoTB
    
    frameGroupProp.ZOrder 0
    frameTBItem.Move frameGroupProp.Left, frameGroupProp.Top
    
    sbCImage_Click 0
        
    UpdateSample True
    FixCtrls4Skin Me
    
    If BelongsToToolbar(GetID, True) > 0 Then
        If IsSubMenu(GetID) Then tsOptions.Tabs.Remove 2
    Else
        tsOptions.Tabs.Remove 2
    End If
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    msgSubClass.SubClassHwnd Me.hwnd, False

End Sub

Private Sub picBackImage_DblClick()

    cmdChange_Click 4

End Sub

Private Sub picLeftNormal_DblClick()

    cmdChange_Click 0

End Sub

Private Sub picLeftOver_DblClick()

    cmdChange_Click 1

End Sub

Private Sub picRightNormal_DblClick()

    cmdChange_Click 2

End Sub

Private Sub picRightOver_DblClick()

    cmdChange_Click 3

End Sub

Private Sub sbApplyOptions_Click()

    PopupMenu frmMain.mnuStyleOptions, , sbApplyOptions.Left, sbApplyOptions.Top + sbApplyOptions.Height

End Sub

Private Sub sbChangeCImage_Click()

    SelImage.FileName = picCorner.tag
    frmRscImages.Show vbModal
    
    With SelImage
        If .IsValid Then
            Select Case SelCIndex
                Case 0
                    MenuGrps(SelId).CornersImages.gcTopLeft = .FileName
                Case 1
                    MenuGrps(SelId).CornersImages.gcTopCenter = .FileName
                Case 2
                    MenuGrps(SelId).CornersImages.gcTopRight = .FileName
                Case 3
                    MenuGrps(SelId).CornersImages.gcLeft = .FileName
                Case 4
                    MenuGrps(SelId).CornersImages.gcRight = .FileName
                Case 5
                    MenuGrps(SelId).CornersImages.gcBottomLeft = .FileName
                Case 6
                    MenuGrps(SelId).CornersImages.gcBottomCenter = .FileName
                Case 7
                    MenuGrps(SelId).CornersImages.gcBottomRight = .FileName
            End Select
            picCorner.tag = .FileName
            picCorner.Picture = LoadPictureRes(picCorner.tag)
        End If
    End With
    
    UpdateCColors
    frmMain.DoLivePreview wbLivePreview, (tsOptions.SelectedItem.key = "tsGroupProperties")

End Sub

Private Sub sbCImage_Click(Index As Integer)

    SelCIndex = Index

    Select Case Index
        Case 0
            lblSelCImage.caption = GetLocalizedStr(716)
            picCorner.tag = MenuGrps(SelId).CornersImages.gcTopLeft
        Case 1
            lblSelCImage.caption = GetLocalizedStr(717)
            picCorner.tag = MenuGrps(SelId).CornersImages.gcTopCenter
        Case 2
            lblSelCImage.caption = GetLocalizedStr(718)
            picCorner.tag = MenuGrps(SelId).CornersImages.gcTopRight
        Case 3
            lblSelCImage.caption = GetLocalizedStr(719)
            picCorner.tag = MenuGrps(SelId).CornersImages.gcLeft
        Case 4
            lblSelCImage.caption = GetLocalizedStr(720)
            picCorner.tag = MenuGrps(SelId).CornersImages.gcRight
        Case 5
            lblSelCImage.caption = GetLocalizedStr(721)
            picCorner.tag = MenuGrps(SelId).CornersImages.gcBottomLeft
        Case 6
            lblSelCImage.caption = GetLocalizedStr(722)
            picCorner.tag = MenuGrps(SelId).CornersImages.gcBottomCenter
        Case 7
            lblSelCImage.caption = GetLocalizedStr(723)
            picCorner.tag = MenuGrps(SelId).CornersImages.gcBottomRight
    End Select
    
    picCorner.Picture = LoadPictureRes(picCorner.tag)
    
    UpdateCColors

End Sub

Private Sub UpdateCColors()

    Dim s As OLE_COLOR
    Dim i As Integer

    For i = 0 To sbCImage.Count - 1
        Select Case i
            Case 0
                s = IIf(LenB(MenuGrps(SelId).CornersImages.gcTopLeft) = 0, &HE0E0E0, &HFF0000)
            Case 1
                s = IIf(LenB(MenuGrps(SelId).CornersImages.gcTopCenter) = 0, &HE0E0E0, &HFF0000)
            Case 2
                s = IIf(LenB(MenuGrps(SelId).CornersImages.gcTopRight) = 0, &HE0E0E0, &HFF0000)
            Case 3
                s = IIf(LenB(MenuGrps(SelId).CornersImages.gcLeft) = 0, &HE0E0E0, &HFF0000)
            Case 4
                s = IIf(LenB(MenuGrps(SelId).CornersImages.gcRight) = 0, &HE0E0E0, &HFF0000)
            Case 5
                s = IIf(LenB(MenuGrps(SelId).CornersImages.gcBottomLeft) = 0, &HE0E0E0, &HFF0000)
            Case 6
                s = IIf(LenB(MenuGrps(SelId).CornersImages.gcBottomCenter) = 0, &HE0E0E0, &HFF0000)
            Case 7
                s = IIf(LenB(MenuGrps(SelId).CornersImages.gcBottomRight) = 0, &HE0E0E0, &HFF0000)
        End Select
        sbCImage(i).BackColor = IIf(i = SelCIndex, IIf(s = &HFF0000, &H800000, &H808080), s)
    Next i
    
    sbRemoveCImage.Enabled = (sbCImage(SelCIndex).BackColor = &H800000)

End Sub

Private Sub sbRemoveCImage_Click()

    Select Case SelCIndex
        Case 0
            MenuGrps(SelId).CornersImages.gcTopLeft = vbNullString
        Case 1
            MenuGrps(SelId).CornersImages.gcTopCenter = vbNullString
        Case 2
            MenuGrps(SelId).CornersImages.gcTopRight = vbNullString
        Case 3
            MenuGrps(SelId).CornersImages.gcLeft = vbNullString
        Case 4
            MenuGrps(SelId).CornersImages.gcRight = vbNullString
        Case 5
            MenuGrps(SelId).CornersImages.gcBottomLeft = vbNullString
        Case 6
            MenuGrps(SelId).CornersImages.gcBottomCenter = vbNullString
        Case 7
            MenuGrps(SelId).CornersImages.gcBottomRight = vbNullString
    End Select
    picCorner.tag = vbNullString
    picCorner.Picture = LoadPicture()
    
    UpdateCColors
    frmMain.DoLivePreview wbLivePreview, (tsOptions.SelectedItem.key = "tsGroupProperties")

End Sub

Private Sub tsOptions_Click()

    Select Case tsOptions.SelectedItem.key
        Case "tsGroupProperties"
            frameGroupProp.Visible = True
            frameTBItem.Visible = False
        Case "tsToolbarItem"
            frameGroupProp.Visible = False
            frameTBItem.Visible = True
    End Select
    
    UpdateSample

End Sub

Private Sub txtLeftH_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtLeftM_Change()
    
    On Error Resume Next
    MenuGrps(SelId).tbiLeftImage.margin = Val(txtLeftM.Text)
    UpdateSample

End Sub

Private Sub txtLeftM_GotFocus()

    SelAll txtLeftM

End Sub

Private Sub txtLeftW_Change()

    On Error Resume Next
    MenuGrps(SelId).tbiLeftImage.w = Val(txtLeftW.Text)
    UpdateSample

End Sub

Private Sub txtLeftW_GotFocus()

    SelAll txtLeftW

End Sub

Private Sub txtLeftW_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtRightH_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtRightM_Change()

    On Error Resume Next
    MenuGrps(SelId).tbiRightImage.margin = Val(txtRightM.Text)
    UpdateSample
    
End Sub

Private Sub txtRightM_GotFocus()

    SelAll txtRightM

End Sub

Private Sub txtRightW_Change()

    On Error Resume Next
    MenuGrps(SelId).tbiRightImage.w = Val(txtRightW.Text)
    UpdateSample

End Sub

Private Sub txtRightW_GotFocus()

    SelAll txtRightW

End Sub

Private Sub txtLeftH_Change()
    
    On Error Resume Next
    MenuGrps(SelId).tbiLeftImage.h = Val(txtLeftH.Text)
    UpdateSample

End Sub

Private Sub txtLeftH_GotFocus()

    SelAll txtLeftH

End Sub

Private Sub txtRightH_Change()

    On Error Resume Next
    MenuGrps(SelId).tbiRightImage.h = Val(txtRightH.Text)
    UpdateSample

End Sub

Private Sub txtRightH_GotFocus()

    SelAll txtRightH

End Sub

Private Sub UpdateSample(Optional IsLoading As Boolean = False)

    If IsUpdating Then Exit Sub
    IsUpdating = True

    With MenuGrps(SelId)
        picLeftNormal.Picture = LoadPictureRes(.tbiLeftImage.NormalImage)
        picLeftOver.Picture = LoadPictureRes(.tbiLeftImage.HoverImage)
        txtLeftW.Text = .tbiLeftImage.w
        txtLeftH.Text = .tbiLeftImage.h
        txtLeftM.Text = .tbiLeftImage.margin
        txtLeftW.Enabled = LenB(.tbiLeftImage.NormalImage) <> 0
        txtLeftH.Enabled = LenB(.tbiLeftImage.NormalImage) <> 0
        txtLeftM.Enabled = LenB(.tbiLeftImage.NormalImage) <> 0
        
        picRightNormal.Picture = LoadPictureRes(.tbiRightImage.NormalImage)
        picRightOver.Picture = LoadPictureRes(.tbiRightImage.HoverImage)
        txtRightW.Text = .tbiRightImage.w
        txtRightH.Text = .tbiRightImage.h
        txtRightM.Text = .tbiRightImage.margin
        txtRightW.Enabled = LenB(.tbiRightImage.NormalImage) <> 0
        txtRightH.Enabled = LenB(.tbiRightImage.NormalImage) <> 0
        txtRightM.Enabled = LenB(.tbiRightImage.NormalImage) <> 0
        
        chkTile.Value = IIf(.tbiBackImage.Tile, vbChecked, vbUnchecked)
        chkAllowCrop.Value = IIf(.tbiBackImage.AllowCrop, vbChecked, vbUnchecked)
        TileImage .tbiBackImage.NormalImage, picBackNormal
        TileImage .tbiBackImage.HoverImage, picBackOver
        
        TileImage .Image, picBackImage
    End With

    If Not IsLoading Then frmMain.DoLivePreview wbLivePreview, (tsOptions.SelectedItem.key = "tsGroupProperties")
    
    IsUpdating = False

End Sub

Private Sub txtRightW_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub

Private Sub LocalizeUI()

    Dim i As Integer
    
    'frameGroupProp.Caption = GetLocalizedStr(124)
    'frameTBItem.Caption = GetLocalizedStr(204)
    tsOptions.Tabs("tsGroupProperties").caption = GetLocalizedStr(124)
    tsOptions.Tabs("tsToolbarItem").caption = GetLocalizedStr(204)
    
    frameNL.caption = GetLocalizedStr(179)
    frameOL.caption = GetLocalizedStr(180)
    frameNR.caption = GetLocalizedStr(179)
    frameOR.caption = GetLocalizedStr(180)

    For i = 0 To cmdChange.Count - 1
        cmdChange(i).caption = GetLocalizedStr(189)
    Next i
    
    sbChangeCImage.caption = GetLocalizedStr(189)
    sbRemoveCImage.caption = GetLocalizedStr(201)
    
    lblImages.caption = GetLocalizedStr(532)
    
    cmdRemoveLeft.caption = GetLocalizedStr(201)
    cmdRemoveRight.caption = GetLocalizedStr(201)
    cmdRemove.caption = GetLocalizedStr(201)
    
    cmdSameLeft.caption = GetLocalizedStr(202)
    cmdSameRight.caption = GetLocalizedStr(202)
    
    lblLeftS.caption = GetLocalizedStr(203)
    lblLeftM.caption = GetLocalizedStr(988)
    lblRightS.caption = GetLocalizedStr(203)
    lblRightM.caption = GetLocalizedStr(988)
    
    lblBackImage.caption = GetLocalizedStr(205)
    
    frmLiveSample.caption = GetLocalizedStr(188)
    
    If lblBackImage.Left < 60 Then lblBackImage.Left = 60
    
    cmdOK.caption = GetLocalizedStr(186)
    cmdCancel.caption = GetLocalizedStr(187)
    
    If Preferences.language <> "eng" Then
        cmdOK.Width = SetCtrlWidth(cmdOK)
        cmdCancel.Width = SetCtrlWidth(cmdCancel)
    End If

End Sub
