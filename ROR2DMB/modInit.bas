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
    
    If FileExists(App.Path + "\ror2dmb.exe.manifest") And Not IsDebug Then
        IsSkinned = InitCommonControlsVB
    End If
    
    frmMain.Show

End Sub


