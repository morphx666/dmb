Attribute VB_Name = "modBevel"
Option Explicit

Public Sub DoBevel(frm As Object, Ctl As Control, ByVal BorderWidth As Integer, ByVal lCorner As Long, ByVal tCorner As Long, ByVal rCorner As Long, ByVal bCorner As Long, Optional ByVal xStyle As Integer, Optional ByVal DoRefresh As Boolean = True)

    Dim AdjustX As Integer
    Dim AdjustY As Integer
    Dim BW As Integer
    Dim tmpC As Long
    
    If DoRefresh Then frm.Refresh
    
    If frm.ScaleMode = vbTwips Then
        AdjustX = Screen.TwipsPerPixelX
        AdjustY = Screen.TwipsPerPixelY
    Else
        AdjustX = 1
        AdjustY = 1
    End If
    
    ' Set the top shading line.
    For BW = 1 To BorderWidth
        frm.CurrentX = Ctl.Left - (AdjustX * BW)
        frm.CurrentY = Ctl.Top - (AdjustY * BW)
        ' Top
        If tCorner >= 0 Then frm.Line -(Ctl.Left + Ctl.Width + (AdjustX * (BW - 1)), Ctl.Top - (AdjustY * BW)), tCorner
        ' Right
        If rCorner >= 0 Then frm.Line -(Ctl.Left + Ctl.Width + (AdjustX * (BW - 1)), Ctl.Top + Ctl.Height + (AdjustY * (BW - 1))), rCorner
        ' Bottom
        If bCorner >= 0 Then frm.Line -(Ctl.Left - (AdjustX * BW), Ctl.Top + Ctl.Height + (AdjustY * (BW - 1))), bCorner
        ' Left
        If lCorner >= 0 Then frm.Line -(Ctl.Left - (AdjustX * BW), Ctl.Top - (AdjustY * BW)), lCorner
        
        ' Check if the border is set to Double
        Select Case xStyle
            Case 1
                If BW = CInt(BorderWidth / 3) Then BW = BW + CInt(BorderWidth / 3)
            Case 2
                If BW = CInt(BorderWidth / 2) Then
                    tmpC = bCorner
                    bCorner = tCorner
                    tCorner = tmpC
                    
                    tmpC = rCorner
                    rCorner = lCorner
                    lCorner = tmpC
                End If
        End Select
    Next BW
    
End Sub

