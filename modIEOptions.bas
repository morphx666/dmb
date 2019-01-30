Attribute VB_Name = "modIEOptions"
Option Explicit

Private IE_ScriptDebugger As String
Private IE_ShowJSErrors As String
Private IE_CDSecurity As String
Private IE_MyComputerSecurity As String

Public Sub ConfigIE()

    On Error Resume Next

    If LenB(IE_ShowJSErrors) = 0 Then
        IE_ScriptDebugger = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "Disable Script Debugger")
        IE_ShowJSErrors = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "Error Dlg Displayed On Every Error")
        IE_CDSecurity = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_LOCALMACHINE_LOCKDOWN\Settings", "LOCALMACHINE_CD_UNLOCK")
        IE_MyComputerSecurity = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_LOCALMACHINE_LOCKDOWN", "iexplore.exe")
    End If

    SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "Disable Script Debugger", "yes"
    SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "Error Dlg Displayed On Every Error", "no"
    SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_LOCALMACHINE_LOCKDOWN\Settings", "LOCALMACHINE_CD_UNLOCK", 1, REG_DWORD
    SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_LOCALMACHINE_LOCKDOWN", "iexplore.exe", 0, REG_DWORD

End Sub

Public Sub RestoreIESettings()

    On Error Resume Next

    If LenB(IE_ShowJSErrors) <> 0 Then
        SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "Disable Script Debugger", IE_ScriptDebugger
        SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "Error Dlg Displayed On Every Error", IE_ShowJSErrors
        SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_LOCALMACHINE_LOCKDOWN\Settings", "LOCALMACHINE_CD_UNLOCK", IE_CDSecurity, REG_DWORD
        SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_LOCALMACHINE_LOCKDOWN", "iexplore.exe", IE_MyComputerSecurity, REG_DWORD
    End If

End Sub

