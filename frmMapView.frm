VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMapView 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Map View"
   ClientHeight    =   2595
   ClientLeft      =   6045
   ClientTop       =   5970
   ClientWidth     =   3255
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   3255
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.TreeView tvMenus 
      Height          =   2175
      Left            =   15
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      WhatsThisHelpID =   20000
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   3836
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   529
      LabelEdit       =   1
      Style           =   7
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "frmMapView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Friend Sub RefreshMap()

    Exit Sub

    Dim g As Integer
    
    tvMenus.Nodes.Clear
    
    For g = 1 To UBound(MenuGrps)
        RenderGroup g
    Next g

End Sub

Private Sub RenderGroup(g As Integer, Optional rc As Node)

    Dim c As Integer
    Dim ng As Node
    Dim nc As Node

    With MenuGrps(g)
        If rc Is Nothing Then
            Set ng = tvMenus.Nodes.Add(, , , .Caption)
        Else
            'Set ng = tvMenus.Nodes.Add(rc, tvwChild, , .Caption)
            Set ng = rc
        End If
        ng.Bold = True
        ng.Expanded = True
        For c = 1 To UBound(MenuCmds)
            With MenuCmds(c)
                If .Parent = g And .Name <> "[SEP]" Then
                    Set nc = tvMenus.Nodes.Add(ng, tvwChild, , .Caption)
                    If .Actions.OnMouseOver.Type = atcCascade Then RenderGroup .Actions.OnMouseOver.TargetMenu, nc
                    If .Actions.OnClick.Type = atcCascade Then RenderGroup .Actions.OnClick.TargetMenu, nc
                    If .Actions.OnDoubleClick.Type = atcCascade Then RenderGroup .Actions.OnDoubleClick.TargetMenu, nc
                End If
            End With
        Next c
    End With

End Sub

Private Sub Form_Resize()

    tvMenus.Move 15, 30, Width - 150, Height - 390

End Sub
