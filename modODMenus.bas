Attribute VB_Name = "modODMenus"
Option Explicit

Public Sub InitMenus(Frm As Form, c As Collection)

    Dim ctrl As Control
    Dim mItem As Menu

    With Frm.Controls("vlmCtrl")
        .UseCustomColors = True
        .MenuBarBackground = Frm.BackColor
        .SetImageList frmMain.ilIcons
        .AutoShowHelp = False
        .ShowTooltip = False
        
        .HighlightStyle = XP
        .DisabledTextColor = &HA5A5A5
        .HighlightedTextColor = &H0&
        .HighlightAppearance = vlFlat
        .MenuBorder = Normal
        .MenuHighlight = &HD1ADAD
        .MenuHighlightBorder = &H800000
        .textColor = &H80000012
        .BitmapBackground = RGB(192, 192, 192)
    End With

    While c.count
        c.Remove 1
    Wend

    For Each ctrl In Frm.Controls
        If TypeOf ctrl Is Menu Then
            Set mItem = ctrl
            With mItem
                If .Caption <> "-" Then
                    c.Add .Name, .Caption
                End If
            End With
        End If
    Next

End Sub
