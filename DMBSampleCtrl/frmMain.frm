VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   1305
   ClientLeft      =   9165
   ClientTop       =   6435
   ClientWidth     =   3735
   LinkTopic       =   "Form1"
   ScaleHeight     =   1305
   ScaleWidth      =   3735
   Begin VB.PictureBox picRsc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   1020
      ScaleHeight     =   435
      ScaleWidth      =   495
      TabIndex        =   1
      Top             =   690
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblDummy 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "DUMMY"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   1575
      TabIndex        =   0
      Top             =   540
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

