Attribute VB_Name = "modInit"
Option Explicit

Private Declare Sub InitCommonControls Lib "Comctl32.dll" ()

Sub Main()

    If FileExists(App.Path + "\dmbwizard.exe.manifest") Then InitCommonControls
    
    frmMain.Show

End Sub
