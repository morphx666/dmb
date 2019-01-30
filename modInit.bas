Attribute VB_Name = "modInit"
Option Explicit

Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Private Const ICC_USEREX_CLASSES = &H200

Public IsSkinned As Boolean

Public Function InitCommonControlsVB() As Boolean
   
   On Error Resume Next
   Dim iccex As tagInitCommonControlsEx
   ' Ensure CC available:
   With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_USEREX_CLASSES
   End With
   InitCommonControlsEx iccex
   InitCommonControlsVB = (Err.number = 0)
   On Error GoTo 0
   
End Function

Sub Main()

    IsDebug = (InStr(1, Command$, "/debug", vbTextCompare)) Or IsInIDE
        
    If FileExists(App.Path + "\dmb.exe.manifest") And Not IsDebug Then
        IsSkinned = InitCommonControlsVB
    End If
    
    frmMain.Show

End Sub

Public Sub ApplyStyle2Ctrl(frm As Form)

    Dim c As Control

    If Not IsSkinned Then Exit Sub
    
    On Error Resume Next
    
    'For Each c In frm.Controls
    '    If c.hWnd <> 0 Then ActivateWindowTheme c.hWnd
    'Next c

End Sub

Public Sub FixCtrls4Skin(frm As Form)

    Dim c As Control
    Dim c1 As Control
    Dim p As Control
    Dim f As Control
    Dim t As Single
    Dim i As Integer
    Dim ho As Integer
    
    If Not IsSkinned Then Exit Sub
    
    On Error Resume Next
    
    For Each f In frm.Controls
        If TypeOf f Is Frame Then
            If Needs2bFixed(frm, f) Then
                For Each c In frm.Controls
                    If c.Container.Name = f.Name Then
                        If Err.number = 0 Then
                            If TypeOf c Is CheckBox Or TypeOf c Is OptionButton Or _
                                (TypeOf c Is Frame And TypeOf c.Container Is Frame) Then
                                i = i + 1
                                Set p = frm.Controls.Add("VB.PictureBox", "dynpic" & i)
                                
                                Set p.Container = f
                                
                                t = 0
                                If f.BorderStyle = 1 Then
                                    t = t + 15 * 7
                                    ho = 15
                                Else
                                    ho = 0
                                End If
                                If f.BorderStyle = 1 And LenB(f.caption) <> 0 Then t = t + 15 * 10
                                p.Left = 15 * 2
                                p.Top = t
                                p.Width = f.Width - p.Left * 2
                                p.Height = f.Height - p.Top * 2 - ho
                                If LenB(f.caption) <> 0 Then p.Height = p.Height + 15 * 16
                                
                                p.BackColor = f.BackColor
                                p.BorderStyle = 0
                                'p.BackColor = RGB(0, 0, i * 100)
                                p.Visible = True
                                
                                For Each c1 In frm.Controls
                                    If Not TypeOf c1 Is Timer Then
                                        If c1.Container.Name = f.Name And c1.Name <> "dynpic" & i Then
                                            Set c1.Container = p
                                            If TypeOf c1 Is Line Then
                                                c1.X1 = c1.X1 - p.Left
                                                c1.X2 = c1.X2 - p.Left
                                                c1.Y1 = c1.Y1 - p.Top
                                                c1.Y2 = c1.Y2 - p.Top
                                            Else
                                                c1.Left = c1.Left - p.Left + 30
                                                c1.Top = c1.Top - p.Top
                                            End If
                                        End If
                                    End If
                                Next c1
                                Exit For
                            End If
                        Else
                            Err.Clear
                        End If
                    End If
                Next c
            End If
        End If
    Next f

End Sub

Private Function Needs2bFixed(frm, f As Frame) As Boolean

    Dim c As Control
    Dim p As Control
    
    On Error Resume Next
    
    For Each c In frm.Controls
        Err.Clear
        Set p = c.Container
        If Err.number = 0 Then
            If p.Name = f.Name Then
                If TypeOf c Is CheckBox Or TypeOf c Is OptionButton Or TypeOf c Is Frame Then
                    If TypeOf c Is Frame Then
                        If (LenB(c.caption) = 0 Or c.BorderStyle = 0) Then
                            Needs2bFixed = Needs2bFixed(frm, c)
                            Exit Function
                        End If
                    End If
                        
                    Needs2bFixed = True
                    Exit Function
                End If
            End If
        End If
    Next c

End Function
