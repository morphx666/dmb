VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private colNames As New Collection
Private colIcons As New Collection
Private pForm As Form

Public Sub Initialize(frm As Form)

    Dim ctrl As Control
    Dim mItem As Menu
    Dim c As String
    Dim idx As String
    
    On Error Resume Next
    
    Set pForm = frm

    With frm.Controls("vlmCtrl")
        .UseCustomColors = True
        .SetImageList frmMain.ilIcons
        .AutoShowHelp = False
        .ShowTooltip = False
        .TootipStyle = vlRectangle
        
        '.MenuBorder = Normal
        '.HighlightAppearance = vlFlat
        
        ' XP Style
'        .HighlightStyle = XP
'        .MenuBackground = RGB(251, 251, 249)
'        .DisabledTextColor = &HA5A5A5
'        .textColor = vbBlack
'        .HighlightedTextColor = vbBlack
'        .MenuHighlight = RGB(191, 202, 217)
'        .MenuHighlightBorder = RGB(113, 150, 240)
'        .BitmapBackground = RGB(229, 229, 222)
        
        ' Windows Default
'        .HighlightStyle = Plain
'        .MenuBackground = frm.BackColor
'        .DisabledTextColor = &H80000011
'        .textColor = &H80000012
'        .HighlightedTextColor = &H8000000E
'        .MenuHighlight = &H8000000D
'        .MenuHighlightBorder = &H8000000D
'        .BitmapBackground = .MenuBackground
        
        ' Inherit from Windows
        '.BitmapBackground = &H80000004
        .BitmapBackground = frm.BackColor
        .DisabledTextColor = &H80000011
        .HighlightAppearance = vlFlat
        .HighlightedTextColor = &H8000000E
        .HighlightStyle = Plain
        .MenuBackground = &H80000004 'menu bar
        .MenuBarBackground = frm.BackColor
        .MenuBorder = Normal
        .MenuHighlight = &H8000000D
        '.MenuHighlightBorder = .MenuHighlight
        .textColor = &H80000007
    End With

    While colNames.Count
        colNames.Remove 1
    Wend
    
    While colIcons.Count
        colIcons.Remove 1
    Wend

    For Each ctrl In frm.Controls
        If TypeOf ctrl Is Menu Then
            Set mItem = ctrl
            With mItem
                idx = ""
                idx = .Index
                If LenB(idx) <> 0 Then idx = "|" + idx
                c = FixCaption(.Caption)
                If LenB(c) <> 0 Then colNames.Add .Name + idx, c
            End With
        End If
    Next

End Sub

Public Property Get IconIndex(ByVal c As String) As Integer

    On Error Resume Next
    
    IconIndex = colIcons(c)

End Property

Public Property Let IconIndex(ByVal c As String, ByVal NewValue As Integer)

    On Error Resume Next
    colIcons.Add NewValue, c

End Property

Public Function Name(ByVal c As String) As String

    On Error Resume Next
    
    If LenB(c) = 0 Then Exit Function
    Name = colNames(FixCaption(c))

End Function

Private Function FixCaption(ByVal c As String) As String

    c = Replace(c, "-", "")
    c = Replace(c, "/", "")
    c = Replace(c, "\", "")
    c = Replace(c, ">", "")
    c = Replace(c, "<", "")
    
    FixCaption = c
    
End Function

Private Sub Class_Terminate()

    Set pForm = Nothing

End Sub
