VERSION 5.00
Begin VB.Form frmMain 
   ClientHeight    =   2430
   ClientLeft      =   5475
   ClientTop       =   5040
   ClientWidth     =   4920
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2430
   ScaleWidth      =   4920
   Begin VB.PictureBox picRsc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   630
      Left            =   2595
      ScaleHeight     =   630
      ScaleWidth      =   720
      TabIndex        =   1
      Top             =   705
      Width           =   720
   End
   Begin VB.Label lblDummy 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "DUMMY"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   540
      TabIndex        =   0
      Top             =   420
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   630
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

